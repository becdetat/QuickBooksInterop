namespace SWXBEN.QuickBooksInterop
{
    public class Customer
    {
        public string Name { get; set; }
        public string Phone { get; set; }
        public Address BillingAddress { get; set; }
        public Address ShippingAddress { get; set; }
        public int? CashSaleCustomerId { get; set; }
        public bool IsCashFlowFinance { get; set; }
    }
}
