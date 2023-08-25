# OpenXML


(Kod OpenXMLWarranty => Form1.cs içinde)
Öncelikle taslak word belgesinde doldurulması gerek alanların önüne Ekle => Yer İşareti sekmesine gelip uygun alanlara isimlendirmelere dikkat edilere yer işaretleri konulmalıdır.
Her alan için farklı isimde yer işareti konulmalıdır.

Projeyi kullanabilmek için documentPath ve newDocumentPath uygun şekilde tanımlanmalıdır.
Aynı zamanda DocumentFormat.OpenXML nuget paketi gereklidir.

Tanımlanan değişkenler örnek garanti belgesi için geçerlidir. Başka belgeler için tanımlamalar değiştirilmelidir.

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


Verilen kısım kodun asıl çalıştığı kısımdır. Her bir değişken için tekrar tekrar yazılması gereken bu kısım bir döngü ve iki dizi ile bir kez yazılarak çalıştırılabilir.
