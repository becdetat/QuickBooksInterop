using System;

namespace SWXBEN.QuickBooksInterop
{
    public class GetInvoiceNumberException : Exception
    {
        public int StatusCode { get; private set; }

        public GetInvoiceNumberException(string message, int statusCode) 
            : base(message)
        {
            StatusCode = statusCode;
        }

        public override string ToString()
        {
            return string.Format(
                "Message: {0}, Status code: {1}",
                Message,
                StatusCode);
        }
    }
}
