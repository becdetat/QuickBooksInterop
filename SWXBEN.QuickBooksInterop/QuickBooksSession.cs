using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using QBFC12Lib;
using swxben.helpers.ObjectExtensions;
using swxben.Windows.Forms;

namespace SWXBEN.QuickBooksInterop
{
    public class QuickBooksSession : IQuickBooksSession
    {
        readonly string _appId;
        readonly string _appName;
        readonly QBSessionManager _sessionManager;
        private readonly Guid _errorRecoveryGuid;
        readonly bool _isOnline;

        public bool IsOnline()
        {
            return _isOnline;
        }

        public QuickBooksSession(string appId, string appName, Guid errorRecoveryId)
        {
            _appId = appId;
            _appName = appName;
            _errorRecoveryGuid = errorRecoveryId;

            try
            {
                _sessionManager = new QBSessionManager();
            }
            catch (Exception ex)
            {
                throw new Exception("Can't create QBSessionManager. Make sure this assembly is compiled for X86 (not x64).", ex);
            }

            try
            {
                _sessionManager.OpenConnection(_appId, _appName);
                _sessionManager.BeginSession("", ENOpenMode.omDontCare);
                _isOnline = true;
            }
            catch (Exception ex)
            {
                ExceptionForm.ShowException("Cannot connect to QuickBooks", ex);
            }
            try
            {
                _sessionManager.ErrorRecoveryID.SetValue("{" + _errorRecoveryGuid.ToString() + "}");
                ErrorRecovery();
            }
            catch (Exception ex)
            {
                ExceptionForm.ShowException("QuickBooks error recovery did not succeed", ex);
            }
        }

        private void ErrorRecovery()
        {
            if (!_isOnline) return;
            if (!_sessionManager.IsErrorRecoveryInfo()) return;

            var errorRecoveryStatus = _sessionManager.GetErrorRecoveryStatus();
            if (errorRecoveryStatus.Attributes.MessageSetStatusCode.Equals("600"))
            {
                MessageBox.Show("The old message set ID does not match any stored IDs and no new message set ID is provided.", "QuickBooks Recovery", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (errorRecoveryStatus.Attributes.MessageSetStatusCode.Equals("9001"))
            {
                MessageBox.Show("Invalid checksum. The new message set ID specified matches the currently stored ID but the checksum fails.", "QuickBooks Recovery", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (errorRecoveryStatus.Attributes.MessageSetStatusCode.Equals("9002"))
            {
                MessageBox.Show("No stored response was found.", "QuickBooks Recovery", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (errorRecoveryStatus.Attributes.MessageSetStatusCode.Equals("9004"))
            {
                MessageBox.Show("Invalid message set ID, greater than 24 characters was given", "QuickBooks Recovery", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (errorRecoveryStatus.Attributes.MessageSetStatusCode.Equals("9005"))
            {
                MessageBox.Show("Unable to store response.", "QuickBooks Recovery", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                var errorRecoveryStatusCode = errorRecoveryStatus.ResponseList.GetAt(0).StatusCode;
                if (errorRecoveryStatusCode == 0)
                {
                    MessageBox.Show("Last request was processed and invoice was added successfully.", "QuickBooks Recovery", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (errorRecoveryStatusCode > 0)
                {
                    MessageBox.Show("There was a warning but the last request was processed successfully.", "QuickBooks Recovery", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Error processing last request", "QuickBooks Recovery", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    var savedMessage = _sessionManager.GetSavedMsgSetRequest();
                    var responseMessageSet = _sessionManager.DoRequests(savedMessage);
                    var response = responseMessageSet.ResponseList.GetAt(0);
                    var statusCode = response.StatusCode;
                    if (statusCode == 0)
                    {
                        var invoiceRet = response.Detail as IInvoiceRet;
                        var resultString = "Following invoice has been successfully submitted to QuickBooks:" + Environment.NewLine + Environment.NewLine;
                        if (invoiceRet != null && invoiceRet.TxnNumber != null)
                        {
                            resultString += "Transaction number = " + invoiceRet.TxnNumber.GetValue().ToString(CultureInfo.InvariantCulture);
                        }
                        MessageBox.Show(resultString, "QuickBooks Recovery", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }

            _sessionManager.ClearErrorRecovery();
            MessageBox.Show("Proceeding with current transaction", "QuickBooks Recovery", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Dispose()
        {
            if (!_isOnline) return;
            _sessionManager.ClearErrorRecovery();
            _sessionManager.EndSession();
            _sessionManager.CloseConnection();
        }

        public string AddInvoice(
            Customer customer,
            DateTime invoiceDate,
            string internalReference,
            IEnumerable<InvoiceLine> invoiceLines,
            string templateRefName)
        {
            if (!_isOnline) return "OFFLINE";
            if (customer == null) throw new ArgumentNullException("customer");

            ErrorRecovery();

            var requestMessageSet = GetMsgSetRequest(_sessionManager);
            requestMessageSet.Attributes.OnError = ENRqOnError.roeContinue;

            var invoiceAdd = requestMessageSet.AppendInvoiceAddRq();

            invoiceAdd.CustomerRef.FullName.SetValue(customer.Name);
            if (!string.IsNullOrEmpty(templateRefName))
            {
                invoiceAdd.TemplateRef.FullName.SetValue(templateRefName);
            }
            invoiceAdd.TxnDate.SetValue(invoiceDate);
            //invoiceAdd.RefNumber.SetValue("xx");//invoiceNumber);
            invoiceAdd.BillAddress.Addr1.SetValue(customer.BillingAddress.Address1);
            invoiceAdd.BillAddress.Addr2.SetValue(customer.BillingAddress.Address2);
            invoiceAdd.BillAddress.Addr3.SetValue(customer.BillingAddress.Address3);
            invoiceAdd.BillAddress.Addr4.SetValue(customer.BillingAddress.Address4);
            invoiceAdd.BillAddress.City.SetValue(customer.BillingAddress.City);
            invoiceAdd.BillAddress.State.SetValue(customer.BillingAddress.State);
            invoiceAdd.BillAddress.PostalCode.SetValue(customer.BillingAddress.PostalCode);
            invoiceAdd.BillAddress.Country.SetValue(customer.BillingAddress.Country);
            invoiceAdd.PONumber.SetValue("");

            foreach (var line in invoiceLines)
            {
                var invoiceLineAdd = invoiceAdd.ORInvoiceLineAddList.Append().InvoiceLineAdd;
                //invoiceLineAdd.ItemRef.FullName.SetValue(line.Name);
                invoiceLineAdd.ServiceDate.SetValue(invoiceDate);
                invoiceLineAdd.ItemRef.FullName.SetValue(line.ItemRef);
                invoiceLineAdd.Desc.SetValue(line.Description);
                invoiceLineAdd.Quantity.SetValue(line.Quantity);
                invoiceLineAdd.ORRatePriceLevel.Rate.SetValue(Math.Round(line.Rate, 2));
                invoiceLineAdd.Amount.SetValue(Math.Round(line.Rate, 2));
            }

            var responseMessageSet = _sessionManager.DoRequests(requestMessageSet);
            var response = responseMessageSet.ResponseList.GetAt(0);

            if (response.StatusCode != 0)
            {
                MessageBox.Show("Unsuccessful response when saving invoice: " + response.StatusCode, "QuickBooks Integration", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new AddInvoiceException(internalReference, customer.Name, response.StatusCode, response.StatusMessage);
            }

            return GetLatestInvoiceNumber(customer.Name, invoiceDate);
        }

        string GetLatestInvoiceNumber(string name, DateTime invoiceDate)
        {
            var requestMessageSet = GetMsgSetRequest(_sessionManager);
            requestMessageSet.Attributes.OnError = ENRqOnError.roeContinue;
            var invoiceQuery = requestMessageSet.AppendInvoiceQueryRq();
            invoiceQuery.IncludeLineItems.SetValue(false);
            invoiceQuery.ORInvoiceQuery.InvoiceFilter.EntityFilter.OREntityFilter.FullNameWithChildren.SetValue(name);
            invoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue(invoiceDate.Date);
            invoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue(invoiceDate.Date.AddDays(1).AddSeconds(-1));

            var responseMessageSet = _sessionManager.DoRequests(requestMessageSet);
            var objResponse = responseMessageSet.ResponseList.GetAt(0);

            if (objResponse.StatusCode == 1)
            {
                throw new GetInvoiceNumberException("No invoices match the query filter used", objResponse.StatusCode);
            }

            var objInvoiceRetList = objResponse.Detail as IInvoiceRetList;

            if (objInvoiceRetList == null || objInvoiceRetList.IsNullOrDefault())
            {
                throw new GetInvoiceNumberException("No invoices returned", -1);
            }

            var invoiceNumbers = new List<int>();
            for (int i = 0; i < objInvoiceRetList.Count; i++)
            {
                var response = objInvoiceRetList.GetAt(i);

                if (response.IsNullOrDefault())
                {
                    throw new GetInvoiceNumberException("Response is null", -1);
                }

                if (response.RefNumber.IsNullOrDefault())
                {
                    continue;
                }

                int n;
                if (int.TryParse(response.RefNumber.GetValue(), out n))
                {
                    invoiceNumbers.Add(n);
                }
            }

            return invoiceNumbers.Any() ? invoiceNumbers.Max().ToString(CultureInfo.InvariantCulture) : "n/a";
        }

        public IEnumerable<Customer> GetCustomers()
        {
            if (!_isOnline) yield break;

            var requestSet = GetMsgSetRequest(_sessionManager);

            Debug.Assert(requestSet != null, "requestSet != null");

            requestSet.AppendCustomerQueryRq();

            var responseSet = _sessionManager.DoRequests(requestSet);
            var response = responseSet.ResponseList.GetAt(0);
            var customerList = response.Detail as ICustomerRetList;

            Debug.Assert(customerList != null, "customerList != null");

            for (var i = 0; i < customerList.Count; i++)
            {
                var customerRet = customerList.GetAt(i);

                yield return GetCustomerFromCustomerRet(customerRet.FullName.GetValue(), customerRet);
            }
        }

        public Customer GetCustomer(string name)
        {
            if (!_isOnline) return new Customer { Name = name };

            var requestSet = GetMsgSetRequest(_sessionManager);

            var customerQuery = requestSet.AppendCustomerQueryRq();
            customerQuery.ORCustomerListQuery.FullNameList.Add(name);

            var responseSet = _sessionManager.DoRequests(requestSet);
            var response = responseSet.ResponseList.GetAt(0);
            var customerList = response.Detail as ICustomerRetList;

            if (customerList == null) return null;

            return customerList.Count == 0 ? null : GetCustomerFromCustomerRet(name, customerList.GetAt(0));
        }

        private static Customer GetCustomerFromCustomerRet(string name, ICustomerRet customerRet)
        {
            var customer = new Customer
                {
                    Name = customerRet.FullName != null ? customerRet.FullName.GetValue() : name,
                    Phone = customerRet.Phone != null ? customerRet.Phone.GetValue() : ""
                };

            if (customerRet.BillAddress != null)
            {
                customer.BillingAddress = GetAddressFromQbAddress(customerRet.BillAddress);
            }
            if (customerRet.ShipAddress != null)
            {
                customer.ShippingAddress = GetAddressFromQbAddress(customerRet.ShipAddress);
            }

            customer.ShippingAddress = customer.ShippingAddress ?? customer.BillingAddress;
            customer.BillingAddress = customer.BillingAddress ?? customer.ShippingAddress;

            customer.ShippingAddress = customer.ShippingAddress ?? Address.Empty;
            customer.BillingAddress = customer.BillingAddress ?? Address.Empty;

            return customer;
        }

        private static Address GetAddressFromQbAddress(IAddress qbAddress)
        {
            var address = Address.Empty;

            if (qbAddress.Addr1 != null) address.Address1 = qbAddress.Addr1.GetValue() ?? "";
            if (qbAddress.Addr2 != null) address.Address2 = qbAddress.Addr2.GetValue() ?? "";
            if (qbAddress.Addr3 != null) address.Address3 = qbAddress.Addr3.GetValue() ?? "";
            if (qbAddress.Addr4 != null) address.Address4 = qbAddress.Addr4.GetValue() ?? "";
            if (qbAddress.City != null) address.City = qbAddress.City.GetValue() ?? "";
            if (qbAddress.State != null) address.State = qbAddress.State.GetValue() ?? "";
            if (qbAddress.PostalCode != null) address.PostalCode = qbAddress.PostalCode.GetValue() ?? "";
            if (qbAddress.Country != null) address.Country = qbAddress.Country.GetValue() ?? "";

            return address;
        }

        static IMsgSetRequest GetMsgSetRequest(QBSessionManager manager)
        {
            short qbXmlMajorVer;
            short qbXmlMinorVer;

            var supportedVersion = QbfcLatestVersion(manager);

            if (supportedVersion >= 6.0)
            {
                qbXmlMajorVer = 6;
                qbXmlMinorVer = 0;
            }
            else if (supportedVersion >= 5.0)
            {
                qbXmlMajorVer = 5;
                qbXmlMinorVer = 0;
            }
            else if (supportedVersion >= 4.0)
            {
                qbXmlMajorVer = 4;
                qbXmlMinorVer = 0;
            }
            else if (supportedVersion >= 3.0)
            {
                qbXmlMajorVer = 3;
                qbXmlMinorVer = 0;
            }
            else if (supportedVersion >= 2.0)
            {
                qbXmlMajorVer = 2;
                qbXmlMinorVer = 0;
            }
            else if (supportedVersion >= 1.1)
            {
                qbXmlMajorVer = 1;
                qbXmlMinorVer = 1;
            }
            else
            {
                qbXmlMajorVer = 1;
                qbXmlMinorVer = 0;
            }

            var requestSet = manager.CreateMsgSetRequest("US", qbXmlMajorVer, qbXmlMinorVer);
            requestSet.Attributes.OnError = ENRqOnError.roeStop;

            return requestSet;
        }
        static double QbfcLatestVersion(QBSessionManager manager)
        {
            var msgset = manager.CreateMsgSetRequest("US", 1, 0);
            msgset.AppendHostQueryRq();

            var queryResponse = manager.DoRequests(msgset);
            var response = queryResponse.ResponseList.GetAt(0);
            var hostResponse = response.Detail as IHostRet;

            Debug.Assert(hostResponse != null, "hostResponse != null");

            var supportedVersions = hostResponse.SupportedQBXMLVersionList;

            int i;
            double lastVersion = 0;

            for (i = 0; i <= supportedVersions.Count - 1; i++)
            {
                var vers = Convert.ToDouble(supportedVersions.GetAt(i));

                if (vers > lastVersion)
                {
                    lastVersion = vers;
                }
            }

            return lastVersion;
        }

        public IDictionary<string, string> GetItemRefs()
        {
            var itemRefs = new Dictionary<string, string>();

            if (_isOnline)
            {
                var requestSet = GetMsgSetRequest(_sessionManager);

                requestSet.AppendItemQueryRq();

                var responseSet = _sessionManager.DoRequests(requestSet);

                var response = responseSet.ResponseList.GetAt(0);
                var itemList = response.Detail as IORItemRetList;


                Debug.Assert(itemList != null, "itemList != null");

                for (var i = 0; i < itemList.Count; i++)
                {
                    var item = itemList.GetAt(i);

                    // currently only service items are returned

                    if (item.ortype == ENORItemRet.orirItemServiceRet)
                    {
                        if (!item.ItemServiceRet.IsActive.GetValue()) continue;

                        var name = item.ItemServiceRet.FullName.GetValue();
                        var account = item.ItemServiceRet.ORSalesPurchase.SalesOrPurchase.AccountRef.FullName.GetValue();
                        itemRefs[name] = account;
                    }
                }
            }

            return itemRefs;
        }
    }
}
