using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;




namespace WordFormatter
{
	class Program
	{
		[STAThread] //to communicate with Windows OS and System dialogs

		static void Main(string[] args)
		{


			//ver 2.0
			if (Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
			{
				MessageBox.Show("Another instance of this application is already running. Please close all instances and try again.");
				return;
			}

			//WordprocessingDocument doc = WordprocessingDocument.Create("C:\\Users\\10254634\\Desktop\\Sample.docx", WordprocessingDocumentType.Document);

			/*	using (doc)
				{
					MainDocumentPart mainPart = doc.AddMainDocumentPart();
					mainPart.Document =  new DocumentFormat.OpenXml.Wordprocessing.Document();
					Body body = mainPart.Document.AppendChild(new Body());
					Paragraph para1 = body.AppendChild(new Paragraph());
					Run run = para1.AppendChild(new Run());
					run.AppendChild(new Text("Titleee"));
				}*/


			//start time
			DateTime StartAt = DateTime.Now;  

			Application.EnableVisualStyles(); //to enable colour fonts and other visual elements on the Form
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new FileSelectForm()); // opens up the dialog
			int numberOfFiles = 0;
			//ver 3.0
			/*try
			{*/
				foreach (String file in FileSelectForm.fileToOpen) //for every file selected by the user
				{

					WordprocessingDocument doc = null;
					String currFile = file;
					numberOfFiles++;
					if (numberOfFiles == 1)
					{
						StartAt = DateTime.Now;
					}
				//ver 2.0 - Handled in Form
				/*if (!Path.GetExtension(currFile).Equals(".docx"))
				{
				MessageBox.Show("Please use an input \".docx\" file. The application will now exit.");
				Environment.Exit(1);
				}*/

				//ver 2.0
				//try
				//{
				doc = WordprocessingDocument.Open(file, true);
					//}
					/*catch (IOException e)
					{
						MessageBox.Show("Please close all the input files and try again.");
						Environment.Exit(1);
					}*/
					MainDocumentPart mainDoc = doc.MainDocumentPart;
					using (doc)
					{						
						
						Body body = mainDoc.Document.Body;


					//Handle empty input .docx file //ver 2.0
					if (body.GetFirstChild<Paragraph>().Descendants().ToList().Count() == 0)
					{
						MessageBox.Show("The input file is empty. The application will now exit.");
						Environment.Exit(1);

					}
					List<Paragraph> initialParas = body.Elements<Paragraph>().ToList();


						//remove the first paragraph and the first table which are not needed
						body.Elements<Table>().First().Remove();

						/*for (int i = 0; i <= 2; i++)
						{*/
						initialParas[0].Remove();
						initialParas[2].Remove();
						/*	}	*/

						mainDoc.Document.Save();
						//All parahraphs directly under body
						List<Paragraph> paragraphs = body.Elements<Paragraph>().ToList();

						//remove the lines between the tables from the whole document
						List<Paragraph> pWithPicture = paragraphs.Where<Paragraph>(p => p.Descendants().OfType<Picture>().ToList().Count() != 0).ToList();
						foreach (var p in pWithPicture)
						{
							p.Remove();
						}

						mainDoc.Document.Save();

						/*var runProp = new RunProperties(
							 new RunFonts()
							 {
								 Ascii = "Arial",
								 ComplexScript = "Arial",
								 HighAnsi = "Arial"
							 }
							 ) ;*/

						//var runFont = new RunFonts { Ascii = "Arial" };
						var fontSize = new FontSize { Val = new StringValue("20") };
						var fontSizeCS = new FontSizeComplexScript { Val = new StringValue("20") };

						/*runProp.Append(fontSize);
						runProp.Append(fontSizeCS);*/

						var runFont = new RunFonts();
						runFont.EastAsia = "Arial";
						runFont.Ascii = "Arial";
						runFont.ComplexScript = "Arial";
						runFont.HighAnsi = "Arial";


						List<Paragraph> pWithRun = new List<Paragraph>();
						List<Paragraph> pWithTextsOutsideTable = new List<Paragraph>();
						List<Paragraph> pWithPprRpr = new List<Paragraph>(); 

						/*try //ver 2.0
						{*/
							List<Paragraph> paragraphsUnderBody = body.Elements<Paragraph>().ToList(); //all ps outside table
							List<Paragraph> allParagraphs = body.Descendants<Paragraph>().ToList(); //includes p in table as well
							List<Table> allTables = body.Elements().OfType<Table>().ToList();
							pWithRun = allParagraphs.Where<Paragraph>(p => p.Descendants().OfType<Run>().ToList().Count() != 0).ToList();
							List<Paragraph> pWithRunOutsideTable = paragraphsUnderBody.Where<Paragraph>(p => p.Descendants().OfType<Run>().ToList().Count() != 0).ToList();
							pWithTextsOutsideTable = pWithRunOutsideTable.Where<Paragraph>(p => p.GetFirstChild<Run>().Descendants().OfType<Text>().ToList().Count() != 0).ToList();
							pWithPprRpr = allParagraphs.Where<Paragraph>(p => p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().ToList().Count() != 0).ToList();
						/*}
						catch(ArgumentNullException e)
						{
							MessageBox.Show("This document has already been formatted and a TestProtocol generated.");
							Environment.Exit(1);
						}*/

						/*foreach (var p in pWithPpr) //outside table - s/w version, test case
						{

							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().GetFirstChild<RunFonts>().Ascii = "Segoe UI"; //add attributes to exiting node
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().GetFirstChild<RunFonts>().ComplexScript = "Arial";
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().GetFirstChild<RunFonts>().HighAnsi = "Arial";
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().Elements().ToList().Append<FontSize>(fontSize); //add a new child
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().Elements().ToList().Add(fontSizeCS);

							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().ChildElements.ToList().Add(fontSize);
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().ChildElements.ToList().Add(fontSizeCS);
							//p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().InsertAfter(fontSize, p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().GetFirstChild<RunFonts>());
							//p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().AppendChild<FontSize>((FontSize)fontSize.CloneNode(true));
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().AppendChild(new FontSize { Val = new StringValue("20") });
							p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>().AppendChild(new FontSizeComplexScript { Val = new StringValue("20") });
							var runs = p.Descendants<Run>().ToList();

							foreach (var r in runs)
							{
								r.GetFirstChild<ParagraphMarkRunProperties>().AppendChild<ParagraphMarkRunProperties>(runProp);
							}

						}
	*/
						foreach (var p in pWithRun) //s/w version, Initial Date and Test Case
						{
							List<Run> runs = p.Descendants<Run>().ToList();

							foreach (var r in runs)
							{
								RunProperties rp = r.GetFirstChild<RunProperties>();
								rp.GetFirstChild<RunFonts>().Ascii = "Calibri";
								rp.GetFirstChild<RunFonts>().ComplexScript = "Calibri";
								rp.GetFirstChild<RunFonts>().HighAnsi = "Calibri";
								rp.AppendChild(new FontSize { Val = new StringValue("22") });
								rp.AppendChild(new FontSizeComplexScript { Val = new StringValue("22") });
							}

						}

						//set font type and size of spacing above software version and initial date and in in table content - heading
						foreach (var p in pWithPprRpr)
						{

							ParagraphMarkRunProperties pMRP = p.GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphMarkRunProperties>();
							pMRP.GetFirstChild<RunFonts>().Ascii = "Calibri";
							pMRP.GetFirstChild<RunFonts>().ComplexScript = "Calibri";
							pMRP.GetFirstChild<RunFonts>().HighAnsi = "Calibri";
							pMRP.AppendChild(new FontSize { Val = new StringValue("22") });
							pMRP.AppendChild(new FontSizeComplexScript { Val = new StringValue("22") });

						}



						//move TestCase line to above Software Version //ver 2.0
						/*List<Paragraph> pWithTestCaseID = pWithTextsOutsideTable.Where<Paragraph>(p => p.GetFirstChild<Run>().GetFirstChild<Text>().Text.Contains("Test case")).ToList();

						foreach (Paragraph p in pWithTestCaseID)
						{
							Paragraph newP = (Paragraph)p.CloneNode(true);
							p.PreviousSibling<Paragraph>().PreviousSibling<Paragraph>().PreviousSibling<Paragraph>().InsertAfterSelf<Paragraph>(newP);
							//p.PreviousSibling<Paragraph>().AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Break());
							p.Remove();
						}*/

						//ver 2.0
						//search and replace some texts
						/*int n = 1;
						foreach (var text in body.Descendants<Text>())
						{
							if (text.Text.Contains("Test case"))
							{
								text.Text = text.Text.Replace("Test case", "1." + n + " ID");
								n++;
							}

							if (text.Text.Contains("Test Instructions"))
							{
								text.Text = text.Text.Replace("Test Instructions", "Action");
								n++;
							}

						}*/

						foreach (var table in body.Elements().OfType<Table>())
						{

							IEnumerable<TableRow> rows = table.Elements<TableRow>();
							//Defining table properties
							TableProperties tblProperties = new TableProperties();

							// Create Table Borders
							TableBorders tblBorders = new TableBorders();

							/*TopBorder topBorder = new TopBorder();
							topBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
							topBorder.Color = "CC0000";
							tblBorders.AppendChild(topBorder);
							BottomBorder bottomBorder = new BottomBorder();
							bottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
							bottomBorder.Color = "CC0000";
							tblBorders.AppendChild(bottomBorder);
							RightBorder rightBorder = new RightBorder();
							rightBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
							rightBorder.Color = "CC0000";
							tblBorders.AppendChild(rightBorder);
							LeftBorder leftBorder = new LeftBorder();
							leftBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
							leftBorder.Color = "CC0000";
							tblBorders.AppendChild(leftBorder);*/

							InsideHorizontalBorder insideHBorder = new InsideHorizontalBorder();
							insideHBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
							insideHBorder.Color = "";
							insideHBorder.Size = 6;
							insideHBorder.Space = 0;
							tblBorders.AppendChild(insideHBorder);
							InsideVerticalBorder insideVBorder = new InsideVerticalBorder();
							insideVBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
							insideVBorder.Color = "808080";
							insideVBorder.Size = 6;
							insideHBorder.Space = 0;
							tblBorders.AppendChild(insideVBorder);

							// Add the table borders to the properties
							//tblProperties.AppendChild(tblBorders);
							// Add the table properties to the table
							table.GetFirstChild<TableProperties>().GetFirstChild<TableBorders>().AppendChild<TableBorders>(tblBorders); //append to existing child 



							// Auto fit at table level and cell level
							TableWidth tblW = new TableWidth();
							tblW.Type = TableWidthUnitValues.Auto;
							tblW.Width = "0";
							table.GetFirstChild<TableProperties>().GetFirstChild<TableWidth>().Remove();
							table.GetFirstChild<TableProperties>().Elements().ToList().Add(tblW);

							var cellsInTable = table.Descendants<TableCell>().ToList();
							foreach (var c in cellsInTable)

							{
								c.GetFirstChild<TableCellProperties>().GetFirstChild<TableCellWidth>().Remove();
								c.GetFirstChild<TableCellProperties>().Elements().ToList().Add(tblW);

							}

							// remove shading in some rows
							var cellsWithShading = cellsInTable.Where<TableCell>(c => c.GetFirstChild<TableCellProperties>().Descendants().OfType<Shading>().ToList().Count() != 0);
							foreach (var c in cellsWithShading)
							{

								c.GetFirstChild<TableCellProperties>().Descendants().OfType<Shading>().ToList()[0].Remove();
							}

							//make column heading bold and alignment correction
							Bold bold = new Bold();
							bold.Val = OnOffValue.FromBoolean(true);
							SpacingBetweenLines spacing = new SpacingBetweenLines();
							spacing.After = "240";
							spacing.Before = "240";

							//to make header row bold
							/*TableRowProperties trpr = new TableRowProperties();
							trpr.AppendChild(new Bold());*/

							table.Elements<TableRow>().ElementAt(1).TableRowProperties.AppendChild<Bold>(new Bold());


							var cellsInHeadingRow = table.GetFirstChild<TableRow>().Descendants().OfType<TableCell>().ToList();
							foreach (var c in cellsInHeadingRow)
							{

								if (c.GetFirstChild<Paragraph>().GetFirstChild<ParagraphProperties>().Descendants().OfType<SpacingBetweenLines>().ToList().Count() != 0)
								{
									c.GetFirstChild<Paragraph>().GetFirstChild<ParagraphProperties>().GetFirstChild<SpacingBetweenLines>().Remove();
								}
								c.GetFirstChild<Paragraph>().GetFirstChild<ParagraphProperties>().Elements().ToList().Add(spacing);
								c.GetFirstChild<Paragraph>().GetFirstChild<ParagraphProperties>().AppendChild(new Justification { Val = JustificationValues.Center });
								c.GetFirstChild<Paragraph>().GetFirstChild<Run>().GetFirstChild<RunProperties>().AppendChild<Bold>(new Bold()); //Bold

							}

							//insert sl no in first cell of every row except heading
							for (int i = 1; i <= rows.ToList().Count() - 1; i++)
							{
								rows.ToList()[i].GetFirstChild<TableCell>().AppendChild(new Paragraph());
								//Setting para properties
								Paragraph newP = rows.ToList()[i].GetFirstChild<TableCell>().Elements<Paragraph>().ElementAt(1);
								newP.AppendChild<ParagraphProperties>(new ParagraphProperties());
								newP.GetFirstChild<ParagraphProperties>().AppendChild(new Justification { Val = JustificationValues.Center });
								//Setting run properties and values
								newP.AppendChild<Run>(new Run());
								Run theRun = newP.GetFirstChild<Run>();
								theRun.AppendChild<RunProperties>(new RunProperties());
								theRun.GetFirstChild<RunProperties>().AppendChild(new RunFonts
								{
									Ascii = "Calibri",
									ComplexScript = "Calibri",
									HighAnsi = "Calibri"
								});
								theRun.GetFirstChild<RunProperties>().AppendChild(new FontSize { Val = new StringValue("22") });
								theRun.GetFirstChild<RunProperties>().AppendChild(new FontSizeComplexScript { Val = new StringValue("22") });
								theRun.AppendChild(new Text(i.ToString()));
							}


						}

						/*Table approvalsTable = body.Elements<Table>().First();
						//Adding Test Lead Name
						TableRow testLeadRow = approvalsTable.Elements<TableRow>().ElementAt(3);
						TableCell testLeadNameCell = testLeadRow.Elements<TableCell>().ElementAt(1);
						testLeadNameCell.AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Text("Kim Davis"));
						//Adding Quality Name
						approvalsTable.Elements<TableRow>().ElementAt(4).Elements<TableCell>().ElementAt(1).AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Text("Sarah Spleidt"));

						int totalNumberOfTablesInDoc = mainDoc.Document.Body.Elements<Table>().Count();
						// Handle only the test case tables
						for (int i = 7; i < totalNumberOfTablesInDoc - 1; i++)
						{

							//doc.MainDocumentPart.Document.Body.AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Text("Software Version: -________________"));
							//doc.MainDocumentPart.Document.Body.AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Break());
							//doc.MainDocumentPart.Document.Body.AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Text("Initial / Date: - ________________"));


							Table table1 = mainDoc.Document.Body.Elements<Table>().ElementAt(i);
							IEnumerable<TableRow> rows = table1.Elements<TableRow>();
							int numberOfCols = rows.ElementAt(0).Elements<TableCell>().Count();
							//Change Heading of 4th column
							rows.ElementAt(0).Elements<TableCell>().ElementAt(3).Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First().Remove();
							rows.ElementAt(0).Elements<TableCell>().ElementAt(3).Elements<Paragraph>().First().Elements<Run>().First().AppendChild(new Text("Pass/Fail"));

							for (int r = 1; r < rows.Count(); r++)

							{
								rows.ElementAt(r).Elements<TableCell>().ElementAt(3).Elements<Paragraph>().First().AppendChild(new Run()).AppendChild(new Text("P____/F____"));

							}

							// Auto fit

							TableCellProperties tableCellProperties = new TableCellProperties();
							for (int r = 0; r < rows.Count(); r++)
							{
								for (int c = 0; c < numberOfCols; c++)
								{
									TableCellWidth tableCellWidth = new TableCellWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
									rows.ElementAt(r).Elements<TableCell>().ElementAt(c).Elements<TableCellProperties>().First().TableCellWidth = tableCellWidth;
								}
							}
						}


						//Adding Footer
						FooterPart footerPart = mainDoc.AddNewPart<FooterPart>();
						string footerPartId = mainDoc.GetIdOfPart(footerPart);
						GenerateFooterPartContent(footerPart);

						IEnumerable<SectionProperties> sections = mainDoc.Document.Body.Elements<SectionProperties>();

						foreach (var section in sections)
						{
							// Delete existing references footers

							section.RemoveAllChildren<FooterReference>();

							// Create new footer reference node

							section.PrependChild<FooterReference>(new FooterReference() { Id = footerPartId });
						}

					}*/

						mainDoc.Document.Save();
					} // end for each doc formatting

					//ver 3.0

					/*mergeFiles(FileSelectForm.templateFileName, currFile, numberOfFiles);*/


				} // all documents formatted
			/*}
			catch(NullReferenceException e) //ver 3.0
			{
				MessageBox.Show("The tool dialog was closed prematurely, please try again.");
				Environment.Exit(1);
			}
*/
			//log the time taken for formatting
			DateTime EndAt = DateTime.Now;
			double timeTaken = Math.Round((EndAt - StartAt).TotalSeconds, 2);
			//int timeTaken = (EndAt - StartAt).Seconds;
			// This will give us the full name path of the executable file including the .exe file:
			string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
			//This will strip just the working path name:
			string strWorkPath = System.IO.Path.GetDirectoryName(strExeFilePath);
			string inLog = "Time taken to format the testcase document: " + timeTaken + " seconds.";
			File.WriteAllText(strWorkPath + "\\" + "log.txt", inLog);
			MessageBox.Show("Formatting Done in " + timeTaken + " seconds.");
			/*public static void GenerateFooterPartContent(FooterPart part)
			{

				Footer footer = new Footer();

				Paragraph footerPara1 = new Paragraph();
				ParagraphProperties footerPara1Properties = new ParagraphProperties();
				ParagraphStyleId objParagraphStyleId = new ParagraphStyleId() { Val = "Footer" };
				Indentation indentation = new Indentation() { Right = "260" };
				footerPara1Properties.Append(objParagraphStyleId);
				footerPara1Properties.Append(indentation);


				Run run01 = new Run();
				Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
				text1.Text = "Becton Dickinson Proprietary Information  ";
				run01.Append(text1);

				Run run02 = new Run();
				Text text2 = new Text();
				text2.Text = "Enter Feature Name";
				run02.Append(text2);

				Run run03 = new Run();
				Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
				text3.Text = "   Rev 1.0";
				run03.Append(text3);



				footerPara1.Append(footerPara1Properties);
				footerPara1.Append(run01);
				footerPara1.Append(run02);
				footerPara1.Append(run03);


				Paragraph footerPara2 = new Paragraph();
				ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties();
				RunStyle runStylePara = new RunStyle() { Val = "PageNumber" };
				paragraphMarkRunProperties.Append(runStylePara);
				FrameProperties frameProperties = new FrameProperties() { Wrap = TextWrappingValues.Around, HorizontalPosition = HorizontalAnchorValues.Margin, VerticalPosition = VerticalAnchorValues.Text, XAlign = HorizontalAlignmentValues.Right, Y = "1" };
				ParagraphProperties footerPara2Properties = new ParagraphProperties();
				ParagraphStyleId objParagraphStyleId2 = new ParagraphStyleId() { Val = "Footer" };
				footerPara2Properties.Append(objParagraphStyleId2);
				footerPara2Properties.Append(frameProperties);
				footerPara2Properties.Append(paragraphMarkRunProperties);


				Run run1 = new Run();
				RunProperties runProperties1 = new RunProperties();
				RunStyle runStyle1 = new RunStyle() { Val = "PageNumber" };
				runProperties1.Append(runStyle1);
				FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
				run1.Append(runProperties1);
				run1.Append(fieldChar1);

				Run run2 = new Run();
				RunProperties runProperties2 = new RunProperties();
				RunStyle runStyle2 = new RunStyle() { Val = "PageNumber" };
				runProperties2.Append(runStyle2);
				FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
				fieldCode1.Text = " PAGE ";
				run2.Append(runProperties2);
				run2.Append(fieldCode1);

				Run run3 = new Run();
				RunProperties runProperties5 = new RunProperties();
				RunStyle runStyle5 = new RunStyle() { Val = "PageNumber" };
				runProperties5.Append(runStyle5);
				FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };
				run3.Append(runProperties5);
				run3.Append(fieldChar3);


				footerPara2.Append(footerPara2Properties);
				footerPara2.Append(run1);
				footerPara2.Append(run2);
				footerPara2.Append(run3);


				footer.Append(footerPara2);
				footer.Append(footerPara1);


				part.Footer = footer;

			} */



		}


		private static void SearchAndReplace(MainDocumentPart mainDocumentPart)
		{
			
				string docText = null;
				using (StreamReader sr = new StreamReader(mainDocumentPart.GetStream()))
				{
					docText = sr.ReadToEnd();
				}

				Regex regexText = new Regex("Test case");
				docText = regexText.Replace(docText, "ID:");

				using (StreamWriter sw = new StreamWriter(mainDocumentPart.GetStream(FileMode.Create)))
				{
					sw.Write(docText);
				}

			mainDocumentPart.Document.Save();
			
		}

		// ver 3.0
		private static void mergeFiles(String templateFile, String currFile, int n)
		{
			String currDir = Path.GetDirectoryName(currFile);
			String fileName = Path.GetFileNameWithoutExtension(currFile);
			String destFilePath = currDir + "\\" + fileName + "_TP.docx";
			//string testFile = @"C:\\Data\\Test.docx";
			//there test two files
			//string[] filepaths = new[] { currDir+"\\T.docx", currFile};
			string[] filepaths = new[] { templateFile, currFile};
			File.Copy(@filepaths[0], destFilePath); //srcFile. DestFile
			//for (int i = 1; i < filepaths.Length; i++)
				//using (WordprocessingDocument myDoc = WordprocessingDocument.Open(@filepaths[0], true))
				using (WordprocessingDocument myDoc = WordprocessingDocument.Open(destFilePath, true))
				{
					MainDocumentPart mainPart = myDoc.MainDocumentPart;
					Body body = mainPart.Document.Body;
					List<Paragraph> pWithRunOutsideTable = body.Elements<Paragraph>().ToList().Where<Paragraph>(p => p.Descendants().OfType<Run>().ToList().Count() != 0).ToList();
					List<Paragraph> pWithTextsOutsideTable = pWithRunOutsideTable.Where<Paragraph>(p => p.GetFirstChild<Run>().Descendants().OfType<Text>().ToList().Count() != 0).ToList();
					Paragraph pWithIntegrationTest = pWithTextsOutsideTable.Where<Paragraph>(p => p.GetFirstChild<Run>().GetFirstChild<Text>().Text.Contains("Integration Test")).ToList()[0];
					string altChunkId = "AltChunkId" + 1;
					AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(
						AlternativeFormatImportPartType.WordprocessingML, altChunkId);
					using (FileStream fileStream = File.Open(@filepaths[1], FileMode.Open))
					{
						chunk.FeedData(fileStream);
					}
					AltChunk altChunk = new AltChunk();
					altChunk.Id = altChunkId;
					//mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.Elements<Paragraph>().Last());
					mainPart.Document.Body.InsertAfter(altChunk, pWithIntegrationTest);
					mainPart.Document.Save();
					myDoc.Close();
				}
		}



	} // end of class



		} //end of namespace

