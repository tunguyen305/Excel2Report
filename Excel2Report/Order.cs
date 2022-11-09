using System;
using System.Collections.Generic;

namespace Report2Pdf
{
	public class Order
	{
		string id;
		public string OrderId
		{
			get { return id; }
			//set { id = value; }
		}

		string userName;
		public string UserName
		{
			get { return userName; }
			set { userName = value; }
		}

		string phone;
		public string Phone
		{
			get { return phone; }
			set { phone = value; }
		}

		string notes;
		public string Notes
		{
			get { return notes; }
			set { notes = value; }
		}

		string addressDetail;
		public string AddrDetail
		{
			get { return addressDetail; }
			set { addressDetail = value; }
		}

		string wards;
		public string Wards
		{
			get { return wards; }
			set { wards = value; }
		}

		string district;
		public string District
		{
			get { return district; }
			set { district = value; }
		}

		string province;
		public string Province
		{
			get { return province; }
			set { province = value; }
		}

		string country;
		public string Country
		{
			get { return country; }
			set { country = value; }
		}

		public string GetFullAddress()
        {
			return string.Format("{0}, {1}, {2}, {3}, {4}.", addressDetail, wards, district, province, country);
        }
		List<Product> lstProducts;
		public List<Product> Products
        {
			get { return lstProducts; }
        }

		public Order(string id)
		{
			this.id = id;
			lstProducts = new List<Product>();

		}


	}
}