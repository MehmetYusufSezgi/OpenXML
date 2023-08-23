﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenXMLWarranty
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void buttonWrite_Click(object sender, EventArgs e)
		{
			string documentPath = @"C:\Users\msezg\OneDrive\Masaüstü\GarantiBelgesiOrnegi.docx";

			string supplierTitle = textBoxSupplierTitle.Text;
			string supplierAddress = textBoxSupplierAddress.Text;
			string supplierPhoneNumber = textBoxSupplierPhoneNumber.Text;
			string supplierFax = textBoxSupplierFax.Text;
			string supplierMail = textBoxSupplierMail.Text;
			string supplierSignature = textBoxSupplierSignature.Text;
			string supplierStamp = textBoxSupplierStamp.Text;

			string sellerTitle = textBoxSellerTitle.Text;
			string sellerAddress = textBoxSellerAddress.Text;
			string sellerPhoneNumver = textBoxSellerPhoneNumber.Text;
			string sellerFax = textBoxSellerFax.Text;
			string sellerMail = textBoxSellerMail.Text;
			string sellerSignature = textBoxSellerSignature.Text;
			string sellerStamp = textBoxSellerStamp.Text;
			string sellerInvoiceDateNAmount = textBoxInvoiceDateNAmount.Text;
			string sellerDeliveryDateNPlace = textBoxDeliveryDateNPlace.Text;

			string productType = textBoxProductType.Text;
			string productBrand = textBoxProductBrand.Text;
			string productModel = textBoxProductModel.Text;
			string warrantyTime = textBoxWarrantyTime.Text;
			string repairTime = textBoxProductRepairTime.Text;
			string bandeloreNSerialNo = textBox.Text;

			string bookmarkSupplierTitle = "bookmarkSupplierTitle";
			string bookmarkSupplierAddress = "bookmarkSupplierAddress";
			string bookmarkSupplierPhone = "bookmarkSupplierPhone";
			string bookmarkSupplierFax = "bookmarkSupplierFax";
			string bookmarkSupplierMail = "bookmarkSupplierMail";
			string bookmarkSupplierSignature = "bookmarkSupplierSignature";
			string bookmarkSupplierStamp = "bookmarkSupplierStamp";

			string bookmarkSellerTitle = "bookmarkSellerTitle";
			string bookmarkSellerAddress = "bookmarkSellerAddress";
			string bookmarkSellerPhone = "bookmarkSellerPhone";
			string bookmarkSellerFax = "bookmarkSellerFax";
			string bookmarkSellerMail = "bookmarkSellerMail";
			string bookmarkSellerSignature = "bookmarkSellerSignature";
			string bookmarkSellerStamp = "bookmarkSellerStamp";
			string bookmarkSellerInvoiceDateNAmount = "bookmarkSellerInvoiceDateNAmount";
			string bookmarkSellerDeliveryDateNPlace = "bookmarkSellerDeliveryDateNPlace";

			string bookmarkProductType = "bookmarkProductType";
			string bookmarkProductBrand = "bookmarkProductBrand";
			string bookmarkProductModel = "bookmarkProductModel";
			string bookmarkProductWarranty = "bookmarkProductWarranty";
			string bookmarkProductRepairDate = "bookmarkProductRepairDate";
			string bookmarkBandoloreNSerialNo = "bookmarkProductBandoloreNSerialNo";

			string[] bookmarks = new string[]
			{
				bookmarkSupplierTitle,
				bookmarkSupplierAddress,
				bookmarkSupplierPhone,
				bookmarkSupplierFax,
				bookmarkSupplierMail,
				bookmarkSupplierSignature,
				bookmarkSupplierStamp,
				bookmarkSellerTitle,
				bookmarkSellerAddress,
				bookmarkSellerPhone,
				bookmarkSellerFax,
				bookmarkSellerMail,
				bookmarkSellerSignature,
				bookmarkSellerStamp,
				bookmarkSellerInvoiceDateNAmount,
				bookmarkSellerDeliveryDateNPlace,
				bookmarkProductType,
				bookmarkProductBrand,
				bookmarkProductModel,
				bookmarkProductWarranty,
				bookmarkProductRepairDate,
				bookmarkBandoloreNSerialNo
			};

			string[] areasInWord = new string[]
			{
				supplierTitle,
				supplierAddress,
				supplierPhoneNumber,
				supplierFax,
				supplierMail,
				supplierSignature,
				supplierStamp,
				sellerTitle,
				sellerAddress,
				sellerPhoneNumver,
				sellerFax,
				sellerMail,
				sellerSignature,
				sellerStamp,
				sellerInvoiceDateNAmount,
				sellerDeliveryDateNPlace,
				productType,
				productBrand,
				productModel,
				warrantyTime,
				repairTime,
				bandeloreNSerialNo
			};


			string newDocumentPath = @"C:\Users\msezg\OneDrive\Masaüstü\NewGarantiBelgesi.docx";
			File.Copy(documentPath, newDocumentPath, true);

			using (WordprocessingDocument document = WordprocessingDocument.Open(newDocumentPath, true))
			{
				for (int i = 0; i < areasInWord.Length; i++)
				{
					BookmarkStart bookmark = document.MainDocumentPart.Document.Body.Descendants<BookmarkStart>()
												  .FirstOrDefault(b => b.Name == bookmarks[i]);
					if (bookmark != null)
					{
						Run run = bookmark.NextSibling<Run>();
						if (run != null)
						{
							Text text = run.GetFirstChild<Text>();
							if (text != null)
							{
								text.Text = areasInWord[i];
							}
						}
					}
				}

				document.Save();
			}

		}
	}
}
