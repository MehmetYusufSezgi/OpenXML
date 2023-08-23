# OpenXML

Projeyi kullanabilmek için documentPath ve newDocumentPath uygun şekilde tanımlanmalıdır.

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
