namespace SWXBEN.QuickBooksInterop
{
    public class InvoiceLine
    {
        public string Name { get; private set; }
        public string Description { get; private set; }
        public double Rate { get; private set; }
        public double Quantity { get; private set; }
        public string ItemRef { get; private set; }

        public InvoiceLine(string name, string description, double rate, double quantity, string itemRef)
        {
            Name = name;
            Description = description;
            Rate = rate;
            Quantity = quantity;
            ItemRef = itemRef;
        }
    }
}
