using System.Text;

namespace SWXBEN.QuickBooksInterop
{
    public class Address
    {
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string Address4 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string PostalCode { get; set; }
        public string Country { get; set; }

        public static Address Empty
        {
            get
            {
                return new Address
                {
                    Address1 = "",
                    Address2 = "",
                    Address3 = "",
                    Address4 = "",
                    City = "",
                    State = "",
                    PostalCode = "",
                    Country = ""
                };
            }
        }

        public string GetFormattedAddress()
        {
            var sb = new StringBuilder();

            if (!string.IsNullOrWhiteSpace(Address1)) sb.AppendLine(Address1);
            if (!string.IsNullOrWhiteSpace(Address2)) sb.AppendLine(Address2);
            if (!string.IsNullOrWhiteSpace(Address3)) sb.AppendLine(Address3);
            if (!string.IsNullOrWhiteSpace(Address4)) sb.AppendLine(Address4);
            if (!string.IsNullOrWhiteSpace(City)) sb.AppendFormat("{0} ", City);
            if (!string.IsNullOrWhiteSpace(State)) sb.AppendFormat("{0} ", State);
            if (!string.IsNullOrWhiteSpace(PostalCode)) sb.AppendFormat("{0} ", PostalCode);
            if (!string.IsNullOrWhiteSpace(Country)) sb.AppendLine().Append(Country);

            return sb.ToString();
        }
    }
}
