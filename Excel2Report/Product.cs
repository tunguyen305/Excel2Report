using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Report2Pdf
{
    public class Product
	{
		string sku;
		public string SKU
		{
			get { return sku; }
			set { sku = value; }
		}

		string name;
		public string Name
		{
			get { return name; }
			set { name = value; }
		}

		string polyphym;
		public string Polyphym
		{
			get { return polyphym; }
			set { polyphym = value; }
		}

		double price;
		public double Price
		{
			get { return price; }
			set { price = value; }
		}

		int count;
		public int Count
        {
			get { return count; }
			set { count = value; }
        }

		public Product()
		{

		}
	}
}
