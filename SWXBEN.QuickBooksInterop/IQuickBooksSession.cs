using System;
using System.Collections.Generic;

namespace SWXBEN.QuickBooksInterop
{
    public interface IQuickBooksSession : IDisposable
    {
        IEnumerable<Customer> GetCustomers();
        Customer GetCustomer(string name);
        string AddInvoice(
            Customer customer,
            DateTime invoiceDate,
            string internalReference,
            IEnumerable<InvoiceLine> invoiceLines,
            string templateRefName);
        IDictionary<string, string> GetItemRefs();
        bool IsOnline();
    }
}
