using System;

namespace SWXBEN.QuickBooksInterop
{
    public class AddInvoiceException : Exception
    {
        public string InternalReference { get; private set; }
        public string CustomerName { get; private set; }
        public int StatusCode { get; private set; }
        public string StatusMessage { get; private set; }

        public AddInvoiceException(string internalReference, string customerName, int statusCode, string statusMessage)
        {
            InternalReference = internalReference;
            CustomerName = customerName;
            StatusCode = statusCode;
            StatusMessage = statusMessage;
        }

        public override string ToString()
        {
            return string.Format(
                "Internal reference: {0}, Customer name: {1}, Status code: {2}, Message: {3}",
                InternalReference,
                CustomerName,
                StatusCode,
                StatusMessage);
        }
    }
}
