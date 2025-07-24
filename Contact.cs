using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LishcaAddIn
{
    public class Contact
    {
        public int id { get; }
        public Rank rank { get; internal set; }
        public string displayName { get; set; }
        public bool pakid { get; set; }
        public Contact[] pkidim { get; set; }
        public List<string> addresses { get; set; }
        public string name { get; private set; }
        public string number { get; private set; }
        public Unit unit { get; private set; }
        public string job { get; private set; }

        public Contact(string address, string displayName)
        {
            this.addresses = new List<string> {address};
            this.displayName = displayName;
            this.name = string.Empty;
            this.number = string.Empty;
            this.unit = null;
            this.rank = null;
        }

        public Contact(int id, Rank rank, string displayName, bool pakid,
            Contact[] pkidim, List<string> addresses, string name, string
            number, Unit unit)
        {
            this.id = id;
            this.rank = rank ?? throw new ArgumentNullException(nameof(rank));
            this.displayName = displayName ?? "Unknown";
            this.pakid = pakid;
            this.pkidim = pkidim;
            if(addresses != null)
            {
                if (addresses.Count > 0)
                {
                    if (addresses[0] != null)
                    {
                        if (addresses[0].Contains(","))
                            this.addresses = addresses[0].Replace(" ",
                                string.Empty).Split(',').ToList();
                        else
                            this.addresses = addresses ?? throw new
                                ArgumentNullException(nameof(addresses));
                    }
                }
            }
            this.name = name ?? string.Empty;
            this.number = number ?? string.Empty;
            this.unit = unit ?? new Unit(string.Empty, string.Empty);
        }

        public Contact()
        {
        }

        public override string ToString()
        {
            if (pakid)
                return "J " + this.name + " [" + "]";
            return this.name;
        }

        public bool SetName(string fullName)
        {
            if (String.IsNullOrWhiteSpace(fullName))
                return false;
            this.name = fullName.Trim();
            return true;
        }

        public bool SetNumber(string phoneNumber)
        {
            if (String.IsNullOrWhiteSpace(phoneNumber) || phoneNumber.Length >
                9)
                return false;
            this.number = phoneNumber.Trim();
            return true;
        }

        public bool SetUnit(Unit unit)
        {
            if (unit is null)
                return false;
            this.unit = unit;
            return true;
        }

        public bool SetRank(Rank rank)
        {
            if (rank is null)
                return false;
            this.rank = rank;
            return true;
        }
    }
}
