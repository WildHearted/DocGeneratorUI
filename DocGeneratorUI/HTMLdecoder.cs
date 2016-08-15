using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HtmlAgilityPack;
using System.Threading.Tasks;
using DocGeneratorCore;

namespace DocGeneratorUI
	{
	/// <summary>
	///      HTML Decoder class is used to instansiate a HTMLdecoder object. The object is used to
	///      decode HTML structure and translate it into Open XML document or Workbook content. Set
	///      the properties begore
	/// </summary>
	class HTMLdecoder
		{
		public enum enumCaptionType
			{
			Image = 1,
			Table = 2
			}

		//++ Object Properties

		/// <summary>
		///      Set the WordProcessing Body immediately after declaring an instance of the
		///      HTMLdecoder object The oXMLencoder requires the WPBody object by reference to add
		///      the decoded HTML to the oXML document.
		/// </summary>
		public Body WPbody { get; set; }

		/// <summary>
		///      The Document Hierarchical Level provides the stating Hierarchical level at which new
		///      content will be added to the document.
		/// </summary>
		public int DocumentHierachyLevel { get; set; }

		/// <summary>
		///      The Additional Hierarchical Level property contains the number of additional levels
		///      that need to be added to the Document Hierarchical Level when processing the HTML
		///      contained in a Enhanced Rich Text column/field.
		/// </summary>
		public int AdditionalHierarchicalLevel { get; set; }

		/// <summary>
		///      ...is used to keep track of the Caption numbers for Tables
		/// </summary>
		public int TableCaptionCounter { get; set; }

		/// <summary>
		///      ... is used to keep track and number the Captions for Images
		/// </summary>
		public int ImageCaptionCounter { get; set; }

		/// <summary>
		///      ... is used to keep track of the Picture IDs that needs to be unique in OpenXML documents.
		/// </summary>
		public int PictureNo { get; set; }

		/// <summary>
		///      The PageWidth property contains the page width of the OXML page into which the
		///      decoded HTML content will be inserted. It is mostly used for image and table
		///      positioning on the page in the OXML document.
		/// </summary>
		public UInt32 PageWidth { get; set; }

		/// The PageHeight property contains the page Height of the OXML page into which the decoded
		/// HTML content will be inserted. It is mostly used for image and table positioning on the
		/// page in the OXML document. </summary>
		public UInt32 PageHeight { get; set; }

		/// <summary>
		/// This property is used to control Content Layer colour coding.
		/// </summary>
		public string ContentLayer { get; set; } = "None";

		/// <summary>
		/// This property will contain the text of the Caption Text to be added after an image
		/// or table
		/// </summary>
		public string CaptionText { get; set; } = string.Empty;

		/// <summary>
		/// This property indicates the type of caption that need to be inserted after an image
		/// or table. It used by the DECODEhtml and ProcessHTMLelements methods to handle
		/// Captions when encoding of OpenXML documents.
		/// </summary>
		public enumCaptionType CaptionType { get; set; } = enumCaptionType.Table;

		/// <summary>
		/// The HyperlinkRelationshipID is used by the DECODEhtml method to handle Hyperlinks in
		/// the encoding of OpenXML documents.
		/// </summary>
		private string HyperlinkImageRelationshipID { get; set; } = string.Empty;

		/// <summary>
		/// The HyperlinkURL contains the ACTUAL hyperlink URL that will be inserted if
		/// required. It used by the DECODEhtml and ProcessHTMLelements methods to to handle
		/// Hyperlinks in the encoding of OpenXML documents.
		/// </summary>
		private string HyperlinkURL { get; set; } = string.Empty;

		/// <summary>
		/// The unique ID of the hyperlink if it need to be inserted. Works in concjunction with
		/// the HyperlinkURL and HyoperlinkImageRelationshipID
		/// </summary>
		private int HyperlinkID { get; set; } = 0;

		/// <summary>
		///      Indicator property that are set once a Hyperlink was inserted for an HTML run
		/// </summary>
		private bool HyperlinkInserted { get; set; } = false;

		private string SharePointSiteURL { get; set; } = string.Empty;

		/// <summary>
		///      Contains the Level of the bullet list
		/// </summary>
		private int BulletLevel { get; set; } = 0;

		/// <summary>
		///      Contains the level of the numbered list
		/// </summary>
		private int NumberedLevel { get; set; } = 0;

		public string ClientName { get; set; } = String.Empty;

		public string WorkString { get; set; } = String.Empty;

		public bool Bold { get; set; } = false;

		public bool Underline { get; set; } = false;

		public bool Italics { get; set; } = false;

		public bool SuperScript { get; set; } = false;

		public bool Subscript { get; set; } = false;

		public bool StrikeTrough { get; set; } = false;

		public WorkTable TableToBeBuild { get; set; }


		//++ Object Methods
		//+----------------
		//+ DecodeHTML
		//+----------------
		/// <summary>
		///      Use this method once a new HTMLdecoder object is initialised and provide at least
		///      all the compulasory(non-optional) properties to the method.
		/// </summary>
		/// <param name="parMainDocumentPart">
		///      [Compulsory] provide the REFERENCE to an already declared MainDocumentPart object of
		///      the document into which the content HTML content needs to be inserted.
		/// </param>
		/// <param name="parDocumentLevel">
		///      [Compulsory] An interger value of 0 to 9 indicating the Heading level or BodyText
		///      Level at which insertion must be inserted.
		/// </param>
		/// <param name="parPageWidthTwips">
		///      [Compulsory] provide the page width into which the content must be inserted. This
		///      needs to be the actual width less margins, gutters and indentations, which and
		///      wherever applicable.
		/// </param>
		/// <param name="parPageHeightTwips">
		///      [Compulsory] provide the page height into which the content must be inserted. This
		///      needs to be the actual height less margins and header/footer offsets, which and
		///      wherever applicable.
		/// </param>
		/// <param name="parHTML2Decode">
		///      [Compulsory] assign the actual HTML string that need to be decoded and inserted to
		///      this parameter.
		/// </param>
		/// <param name="parTableCaptionCounter">
		///      [Compulsory] provide the REFERENCE to an integer containing the current
		///      TableCaptionCounter that need to be used if the HTML contains a table.
		/// </param>
		/// <param name="parImageCaptionCounter">
		///      [Compulsory] provide the REFERENCE to an integer containing the current
		///      ImageCaptionCounter that need to be used if the HTML contains a table.
		/// </param>
		/// <param name="parPictureNo">
		///      [Compulsory] provide the REFERENCE to an integer containing the current Picture
		///      Number/Counter that need to be used if the HTML contains a imagetable.
		/// </param>
		/// <param name="parHyperlinkID">
		///      [Compulsory] hyperlinks needs to be unique in the OpenXML document, therefore a
		///      unique value must ALWAYS be inserted into the OpenXML document else it becomes invalid.
		/// </param>
		/// <param name="parHyperlinkURL">
		///      [Compulsory] but can be null if no hyperlink needs to be inserted referencing the
		///      complete HTML section.
		/// </param>
		/// <param name="parHyperlinkImageRelationshipID">
		///      [Compulsory] but can be null, if the parHyperlinkURL is null.this is the
		///      Relationship id, provided used in the OPENXml document where the image to which the
		///      hyperlink in parHyperlink and parHyperlinkID needs to be linked.
		/// </param>
		/// <param name="parContentLayer">
		///      [optional] defaults to "None" else it must be one of the following string values:
		///      "Layer1", "Layer2" or "Layer3"
		/// </param>
		/// <returns>
		///      returns a boolean value of TRUE if insert was successfull and FALSE if there was any
		///      form of failure during the insertion.
		/// </returns>
		public bool DecodeHTML(
			ref MainDocumentPart parMainDocumentPart,
			int parDocumentLevel,
			string parHTML2Decode,
			ref int parTableCaptionCounter,
			ref int parImageCaptionCounter,
			ref int parPictureNo,
			ref int parHyperlinkID,
			string parSharePointSiteURL,
			string parHyperlinkURL = "",
			string parHyperlinkImageRelationshipID = "",
			string parContentLayer = "None")

			{

			//+Update the Properties of the object...
			this.SharePointSiteURL = parSharePointSiteURL;
			this.DocumentHierachyLevel = parDocumentLevel;
			this.AdditionalHierarchicalLevel = 0;
			this.TableCaptionCounter = parTableCaptionCounter;
			this.ImageCaptionCounter = parImageCaptionCounter;
			this.PictureNo = parPictureNo;
			this.HyperlinkImageRelationshipID = parHyperlinkImageRelationshipID;
			this.HyperlinkURL = parHyperlinkURL;
			this.ContentLayer = parContentLayer;
			this.HyperlinkID = parHyperlinkID;

			//+Variables:
			//- The sequence of the values in the tuple variable is: 
				//-**1** *Bullet Levels*, 
				//-**2** *Number Levels* 
			Tuple<int, int> bulletNumberLevels = new Tuple<int, int>(0, 0);
			int bulletLevel = 0;
			int numberLevel = 0;
			int headingLevel = 0;
			int tableWidth = 0;
			int columnWidth = 0;
			bool isHeaderRow = false;
			string tableAttrValue = String.Empty;
			int tableRowSpanQty = 0;
			int tableColumnSpanQty = 0;
			bool newParagraph = true;
			bool isDIVempty = true;

			Paragraph objNewParagraph = new Paragraph();
			Run objRun = new Run();

			try
				{

				//-Declare the htmlData object instance
				HtmlAgilityPack.HtmlDocument htmlData = new HtmlAgilityPack.HtmlDocument();
				//-Load the HTML content into the htmlData object.
				//-Load a string into the HTMLDocument
				
				htmlData.LoadHtml(parHTML2Decode);
				Console.WriteLine("HTML: {0}", htmlData.DocumentNode.OuterHtml);
				//- Set the ROOT of the data loaded in the htmlData
				var htmlRoot = htmlData.DocumentNode;
				Console.WriteLine("Root Node Tag..............: {0}", htmlRoot.Name);
				Console.WriteLine("______________________________HTML decoding iterations begin ______________________________________");
				foreach(HtmlNode node in htmlData.DocumentNode.DescendantsAndSelf())
					{

					//Console.WriteLine(">-- {0} --<", node.Name);

					switch(node.Name)
						{
						//+<DIV>
						case "div":
							{
							//-Check for other tags in the **div** tag
							isDIVempty = true;
							if(node.HasChildNodes)
								{
								//-Check if there are any other valid HTML closing tags succeeding the **<DIV>**

								foreach(HtmlNode decendentNode in node.Descendants())
									{
									//-Check for any valid content that need to be processed...
									if(decendentNode.Name == "#text"   //-//text
									|| decendentNode.Name == "p"       //-//paragraph
									|| decendentNode.Name == "ol"     //-//Organised List
									|| decendentNode.Name == "ul"     //-//Unorganised List
									|| decendentNode.Name == "li"      //-//ListItem
									|| node.XPath.Contains("/img")     //-//Image
									|| node.XPath.Contains("/table")   //-//Table
									|| decendentNode.Name.Contains("h")//-//Heading
									)
										{
										//- Just CONSUME the div tag and process the tags in the subsequent node cycles...
										isDIVempty = false;
										break;
										}
									}
								}
							
							if(isDIVempty)
								{
								//-When the code reach this point it means that there is text in an isolated **DIV** tag
								//-which means the text cannot be properly formatted because it doesn't contain valid...
								//-Just check first if there is not a populated paragraph that need to be inserted into the document before the content error is raised
								if(objNewParagraph != null && !String.IsNullOrEmpty(objNewParagraph.InnerText))
									{ //-Write the *paragraph* to the document...
									this.WPbody.Append(objNewParagraph);
									objNewParagraph = null;
									}
								//-HTML tags - **THEREFORE** a format issue is raised...
								throw new InvalidContentFormatException("The content was probably pasted from an external source without "
									+ "formatting in afterwards. "
									+ "Please inspect the content and resolve the issue, by formatting the relevant content with the "
									+ "Enhanced Rich Text styles. |" + node.InnerText + "|");
								}
							else
								{
								/*this.WorkString = CleanText(node.InnerText, this.ClientName);
								if(this.WorkString != String.Empty)
									{
									Console.Write(" <{0}>|{1}|", node.Name, this.WorkString);
									if(this.WorkString != string.Empty)
										{
										objNewParagraph = oxmlDocument.Construct_Paragraph(
											parBodyTextLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: this.WorkString,
											parContentLayer: this.ContentLayer);
										objNewParagraph.Append(objRun);
										this.WPbody.Append(objNewParagraph);
										}
									}
									*/
								}
							break;
							}

						//+<#text> **TEXT**
						case "#text":
							{
							//-First, clean the string by removing unwanted and unnecessary characters
							this.WorkString = CleanText(node.InnerText, this.ClientName);
							//-Determine if the text belongs in a **Table**
							if(node.XPath.Contains("/table"))
								{//-Text is part of a table...
								//-Determine if the type of table cell...
								if(node.XPath.Contains("/th"))		//-**First / Heading** Cell
									{
									//!Create a Header Cell
									}
								
								else if(node.XPath.Contains("/td"))	//-**Normal** Cell
									{
									//!Create a Normal Cell

									}
								}
							else      //-The text is **NOT** part of a *Table*...
								{
								//!Clear the table Cell object...
								}


							if(this.WorkString != String.Empty)
								{
								if(node.XPath.Contains("/strong"))
									{
									Console.Write("|BOLD|");
									this.Bold = true;
									}
								else
									this.Bold = false;

								if(node.XPath.Contains("/em"))
									{
									Console.Write("|ITALICS|");
									this.Italics = true;
									}
								else
									this.Italics = false;

								if(node.XPath.Contains("/span"))
									if(node.ParentNode.HasAttributes)
										{
										foreach(HtmlAttribute attributeItem in node.ParentNode.Attributes)
											{
											if(attributeItem.Value.Contains("underline"))
												{
												this.Underline = true;
												Console.Write("|UNDELINE|");
												}
											else
												this.Underline = false;
											}
										}
									else
										this.Underline = false;
								else
									this.Underline = false;

								if(node.XPath.Contains("/sub"))
									{
									Console.Write("|SUBSCRIPT|");
									this.Subscript = true;
									}
								else
									this.Subscript = false;

								if(node.XPath.Contains("/sup"))
									{
									Console.Write("|SUPERSCRIPT|");
									this.SuperScript = true;
									}
								else
									this.SuperScript = false;

								//-Insert the **text** if the string is *not* empty
								if(!String.IsNullOrWhiteSpace(this.WorkString))
									{
									Console.Write("[{0}]", this.WorkString);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: this.WorkString,
										parContentLayer: this.ContentLayer,
										parBold: this.Bold,
										parItalic: this.Italics,
										parUnderline: this.Underline,
										parSubscript: this.Subscript,
										parSuperscript: this.SuperScript);

									// Check if a hyperlink must be inserted
									if(this.HyperlinkImageRelationshipID != "" && this.HyperlinkURL != "")
										{
										if(this.HyperlinkInserted == false)
											{
											this.HyperlinkID += 1;
											DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing =
												oxmlDocument.ConstructClickLinkHyperlink(
												parMainDocumentPart: ref parMainDocumentPart,
												parImageRelationshipId: this.HyperlinkImageRelationshipID,
												parClickLinkURL: this.HyperlinkURL,
												parHyperlinkID: this.HyperlinkID);
											objRun.Append(objDrawing);
											this.HyperlinkInserted = true;
											}
										}
									objNewParagraph.Append(objRun);
									}
								}
							break;
							}

						//+<p> - **Paragraph**
						case "p":
							{
							//-Check if the **objNewParagraph** is *NOT* null **AND** that it actually contains text
							if(objNewParagraph != null && !String.IsNullOrEmpty(objNewParagraph.InnerText))
								{ //-Write the *paragraph* to the document...
								this.WPbody.Append(objNewParagraph);
								objNewParagraph = null;
								}

							//-Check if the paragraph is part of a bullet- number- list
							if(node.XPath.Contains("/ol") || node.XPath.Contains("/ul") && node.XPath.Contains("/li"))
								{ //-If it is :. get the number of bullet - and number - levels in the xPath
								bulletNumberLevels = GetBulletNumberLevels(node.XPath);
								bulletLevel = bulletNumberLevels.Item1;
								numberLevel = bulletNumberLevels.Item2;
								//- now exit the loop, bto process the **"#text"** or other child tags...
								break;
								}
							else
								{
								bulletLevel = 0;
								numberLevel = 0;
								}
							
							//-Check if a **NEW** paragraph must be initialised or whether the existing paragraph needs to be used to add run text.
							if(newParagraph)
								objNewParagraph = new Paragraph();

							if(node.HasChildNodes)
								{
								Console.Write("\n\t <{0}> ", node.Name);
								}
							else
								{//?Does this means it is an empty paragraph?
								this.WorkString = node.InnerText;
								if(this.WorkString != String.Empty)
									Console.Write(" <{0}>", node.Name);
								}
							break;
							}
					
						//+<H?> **Heading 1-4**
						case "h1":
						case "h2":
						case "h3":
						case "h4":
							{
							//-Get the *Heading level* and set the **headingLevel** value
							if(!int.TryParse(node.Name.Substring(1, (node.Name.Length - 1)), out headingLevel))
								{ headingLevel = 0; }
							Console.Write("\n {0} + <{1}>", new String('\t', headingLevel * 2), node.Name);
							//- Set the **this.AdditionalHierarchicalLevel** to the headingLevel value
							this.AdditionalHierarchicalLevel = headingLevel;

							//-Check if there is a populated paragraph
							if(objNewParagraph != null 
							&& objNewParagraph.InnerText != null)
								{ //-Write the *paragraph* to the document...
								this.WPbody.Append(objNewParagraph);
								}
							//-Create the paragrpah for the **Heading**
							objNewParagraph = oxmlDocument.Construct_Heading(
								parHeadingLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);

							//-if there are no child nodes, check if the innterText is also blank
							if(!node.HasChildNodes)
								{
								if(node.InnerText == String.Empty)
									{//-Destroy the paragraph because it will be a blank heading in the document...
									objNewParagraph = null;
									}
								}
							break;
							}
						//+ <UL> **Unorganised List**
						case "ul":
							{
							//-Check if there is a populated paragraph
							if(objNewParagraph != null
							&& objNewParagraph.InnerText != null)
								{ //-Write the *paragraph* to the document...
								this.WPbody.Append(objNewParagraph);
								objNewParagraph = null;
								}
							else
								objNewParagraph = null;

							if(node.HasChildNodes)
								{
								//Console.Write("\n {0} <{1}>", Tabs(headingLevel) + Tabs(bulletLevel), node.Name);
								}
							else
								{
								//?Don't think the code will ever reach here...
								this.WorkString = node.InnerText;
								Console.WriteLine("\t\t\t <{0}>|{1}|", node.Name, this.WorkString);
								}
							break;
							}

						//+<ol> **Organised List**
						case "ol":
							{
							//-Check if there is a populated paragraph that **CONTAINS** text...
							if(objNewParagraph != null
							&& !String.IsNullOrWhiteSpace(objNewParagraph.InnerText))
								{ //-Write the *paragraph* to the document...
								this.WPbody.Append(objNewParagraph);
								objNewParagraph = null;
								}
							else
								objNewParagraph = null;

							if(node.HasChildNodes)
								{
								//Console.Write("\n {0} <{1}>", Tabs(headingLevel) + Tabs(bulletLevel), node.Name);
								}
							else
								{
								//?Don't think the code will ever reach here...
								this.WorkString = node.InnerText;
								Console.WriteLine("\t\t\t <{0}>|{1}|", node.Name, this.WorkString);
								}
							break;
							}

						//+<li>  **List Item**
						case "li":
							{
							//-Determine the number of bullet- and number- levsel from the xPath
							bulletNumberLevels = GetBulletNumberLevels(node.XPath);
							bulletLevel = bulletNumberLevels.Item1;
							numberLevel = bulletNumberLevels.Item2;

							//-Check if there is a populated paragraph that contains text and write it to the Document before initiating a new paragraph...
							if(objNewParagraph != null
							&& !String.IsNullOrEmpty(objNewParagraph.InnerText))
								{ //-Write the *paragraph* to the document...
								this.WPbody.Append(objNewParagraph);
								objNewParagraph = null;
								}

							//-Construct the paragraph with the bullet or number :. depends on the value of the bulletLevel...
							if(bulletLevel > 0)
								{//- if it is a **Bullet** list entry, create a new **Pargraph** *Bullet* object...
								objNewParagraph = oxmlDocument.Construct_BulletNumberParagraph(
									parIsBullet: true,
									parBulletLevel: bulletLevel);
								}
							//- check if it is **Organised/Number list** item
							else if(numberLevel > 0)
								{//-if it is a **Number** list entry, create a new **Pargraph** *Number* object instance...
									objNewParagraph = oxmlDocument.Construct_BulletNumberParagraph(
									parIsBullet: false,
									parBulletLevel: numberLevel);
								}
							else
								{
								//?condition should never materialise, unless the bullet/number
								break;
								}

							if(node.HasChildNodes)
								{
								if(bulletLevel > 0)
									{
									Console.Write("\n {0} - <{1}>", new String('\t',headingLevel + bulletLevel), node.Name);
									}
								else if(numberLevel > 0)
									{
									Console.Write("\n {0} {1}. <{2}>", new String('\t',(headingLevel  + numberLevel)), numberLevel, node.Name);
									}
								}
							else
								{
								//?Check if thia code is ever reached...
								this.WorkString = node.InnerText;
								Console.WriteLine("\t\t\t <{0}>|{1}|", node.Name, this.WorkString);
								}
							break;
							}

						//++Image
						case "img":
							{


							break;
							}

						//++Table
						case "table":
							{
							//!Add a check that prevent CASCADING tables...

							Console.Write("\n\n <Table> ");
							tableWidth = 0;
							//-Check if the table has **attributes** define to obtain the table's **width**...
							if(!node.HasAttributes)  //- The table doesn't have attributes 
								{
								Console.WriteLine("\n ERROR - No attributes defined for the table");
								//-HTML tags - **THEREFORE** a format issue is raised...
								throw new InvalidContentFormatException("The TABLE's width is missing, therefore the table cannot be inserted into the "
									+ "document. Please inspect the content and resolve the issue, by formatting the relevant content with the "
									+ "Enhanced Rich Text styles.");
								}

							else 
								{
								//-Define the Table Instance 
								this.TableToBeBuild = new WorkTable();
								this.TableToBeBuild.Active = true;      //-Set Table Mode to Active
								this.TableToBeBuild.GridDone = false;   //-The table's **GRID** is not yet determined...

								//-Process the table attributes to determine how to format the table...
								foreach(HtmlAttribute tableAttr in node.Attributes)
									{
									switch(tableAttr.Name)
										{
										//-use the **summary** to obtain and set the **Table Caption**
										case "summary": //- get the table caption
											{
											if(tableAttr.Value == null)
												this.TableToBeBuild.Caption = String.Empty;
											else
												this.TableToBeBuild.Caption = tableAttr.Value;
											break;
											}

										//-get the table **width** from the style attribute
										case "style":
											{
											//- Check that the style contains the table width as part of the style
											if(tableAttr.Value.Contains("width:"))
												{
												if(tableAttr.Value.Contains("%"))
													{
													if(int.TryParse(tableAttr.Value.Substring(
														tableAttr.Value.IndexOf(":") + 2,
														(tableAttr.Value.IndexOf("%") - tableAttr.Value.IndexOf(":") - 2)),
														out tableWidth))
														{ //-Successfully parsed the integer...
														this.TableToBeBuild.OriginalTableWidthType = WorkTable.enumWidthType.Percentage;
														this.TableToBeBuild.OriginalTableWidthValue = tableWidth;
														}
													}
												else //-Table width is **NOT** a percentage :. px value
													{
													if(int.TryParse(tableAttr.Value.Substring(
														tableAttr.Value.IndexOf(":") + 2,
														(tableAttr.Value.IndexOf("px") - tableAttr.Value.IndexOf(":") - 2)),
														out tableWidth))
														{
														//-Successfully parsed the integer...
														this.TableToBeBuild.OriginalTableWidthType = WorkTable.enumWidthType.Pixel;
														this.TableToBeBuild.OriginalTableWidthValue = tableWidth;
														}
													}
												}
											break;
											}

										//-The table Width is specified as an attribute
										case "width":
											{
											if(tableAttr.Value.Contains("%"))
												{
												if(int.TryParse(tableAttr.Value.Substring(
													tableAttr.Value.IndexOf(":") + 2,
													(tableAttr.Value.IndexOf("%") - tableAttr.Value.IndexOf(":") - 2)),
													out tableWidth))
													{
													//-Successfully parsed the tableWidth...
													this.TableToBeBuild.OriginalTableWidthType = WorkTable.enumWidthType.Percentage;
													this.TableToBeBuild.OriginalTableWidthValue = tableWidth;
													}
												}
											else //-Table width is **NOT** a percentage :. px value
												{
												if(int.TryParse(tableAttr.Value.Substring(
													tableAttr.Value.IndexOf(":") + 2,
													(tableAttr.Value.IndexOf("px") - tableAttr.Value.IndexOf(":") - 2)),
													out tableWidth))
													{
													//-Successfully parsed the integer...
													this.TableToBeBuild.OriginalTableWidthType = WorkTable.enumWidthType.Pixel;
													this.TableToBeBuild.OriginalTableWidthValue = tableWidth;
													}
												}
											break; //break out of the SWITCH
											}
										}
									}

								}

							if(this.TableToBeBuild.OriginalTableWidthValue > 0)
								{
								Console.Write("\t ...Width: {0} {1}...", this.TableToBeBuild.OriginalTableWidthValue, this.TableToBeBuild.OriginalTableWidthType);
								}
							else
								{
								Console.WriteLine("\n ERROR - No attributes defined for the table");
								//-HTML tags - **THEREFORE** a format issue is raised...
								throw new InvalidContentFormatException("The TABLE's width is missing, therefore the table cannot be inserted into the "
									+ "document. Please inspect the content and resolve the issue, by formatting the relevant content with the "
									+ "Enhanced Rich Text styles.");
								}
							
							if(this.TableToBeBuild.OriginalTableWidthType == WorkTable.enumWidthType.Percentage)
								{
								this.TableToBeBuild.WidthPrecetage = this.TableToBeBuild.OriginalTableWidthValue;
								this.TableToBeBuild.WidthPixels = Convert.ToInt16((this.PageWidth * this.TableToBeBuild.OriginalTableWidthValue) / 100);
								}
							else //-OriginalTableWidthType is **Pixels**
								{
								this.TableToBeBuild.WidthPrecetage = Convert.ToInt16((this.TableToBeBuild.OriginalTableWidthValue / this.PageWidth) * 100));
								this.TableToBeBuild.WidthPixels = this.TableToBeBuild.OriginalTableWidthValue;
								}

							//!Build the Table Grid...
							

							break;
							}
						//+Table Body = **<tb>**
						case "tbody":
							{
							//-Just ignore it.
							break;
							}


						//+Table Row = **<tr>**
						case "tr":
							{
							Console.Write("\n\t <Row>");

							//Determine if it is a **HEADER** row
							if(node.HasAttributes)  //- The table doesn't have attributes 
								{
								//-Determine if the row is a **HEADER** row...
								foreach(HtmlAgilityPack.HtmlAttribute tableAttr in node.Attributes)
									{
									//use the **class** to dertermine whether the row is specified as a *TableHeader*
									if(tableAttr.Name == "class")
										{
										if(tableAttr.Value == null)
											isHeaderRow = false;
										else
											{
											if(tableAttr.Value.Contains("HeaderRow"))
												isHeaderRow = true;
											else
												isHeaderRow = false;
											}
										}
									}
								}
							else //- The table doen't have a Header row
								{ //-Therefore it will also not have a Header Row
								isHeaderRow = false;
								}

							Console.Write(" ... Is it a Header Row? {0}", isHeaderRow);
							break;
							}

						//+<th> - **Table Header Cell**
						case "th":
							{
							//-If the **table header** was not determined in the class of the table row *<tr>* tag, then the table has a header row if a *<th>* tag is present
							//-therefore set the **isHeaderRow** to *true*.
							isHeaderRow = true; //-Determined the table has a **HEADER** row
							if(node.HasAttributes)  //- The table Header Cell has attributes 
								{
								
								tableAttrValue = String.Empty;
								columnWidth = 0;
								//-get the **Row Span** if it is defined...
								tableAttrValue = node.Attributes.Where(a => a.Name == "rowspan").Single().Value.ToString();
								//-if no **rowspan** is found, set the value to 1
								if(String.IsNullOrWhiteSpace(tableAttrValue))
									tableRowSpanQty = 1;
								else
									{
									if(!int.TryParse(tableAttrValue, out tableRowSpanQty))
										tableRowSpanQty = 1;
									}

								//-Get the **Column Span** if there is a value
								tableAttrValue = String.Empty;
								tableAttrValue = node.Attributes.Where(a => a.Name == "colspan").Single().Value.ToString();
								if(String.IsNullOrWhiteSpace(tableAttrValue))
									tableColumnSpanQty = 1;
								else
									{
									if(!int.TryParse(tableAttrValue, out tableColumnSpanQty))
										tableColumnSpanQty = 1;
									}

								//-Get the column **Width** if specified
								tableAttrValue = node.Attributes.Where(a => a.Name == "style" && a.Value.Contains("width")).Single().Value.ToString();
								if(tableAttrValue == String.Empty)
									columnWidth = 0;
								else
									{
									columnWidth = 0;
									if(tableAttrValue.Contains("%"))
										{
										if(int.TryParse(tableAttrValue.Substring(
											tableAttrValue.IndexOf(":") + 2,
											(tableAttrValue.IndexOf("%") - tableAttrValue.IndexOf(":") - 2)),
											out columnWidth) == false)
											{
											//- Could not parse the integer which means the tableWidth remains as it was before the parse.
											}
										}
									else //-Column width is **NOT** a percentage :. px value
										{
										if(int.TryParse(tableAttrValue.Substring(
											tableAttrValue.IndexOf(":") + 2,
											(tableAttrValue.IndexOf("px") - tableAttrValue.IndexOf(":") - 2)),
											out tableWidth) == false)
											{
											//-Could not parse the integer which means the tableWidth remains as it was before the parse.
											}
										}
									}
								}
							else //- the table Header Cell doesn't have any attributes
								{
								tableColumnSpanQty = 1;
								tableRowSpanQty = 1;
								columnWidth = 0;
								}
							Console.Write("\n\t\t<col> Header Column: Width: {0}\t RowSpan: {1}\t ColumnSpan: {2}\t", columnWidth, tableRowSpanQty, tableColumnSpanQty);
							break;
							}
						//+TableCell - **<td>**
						case "td":
							{
							//-If the **table header** was not determined in the class of the table row *<tr>* tag, then the table has a header row if a *<th>* tag is present
							isHeaderRow = false;

							if(node.HasAttributes)  //- The table Header Cell has attributes 
								{
								foreach(HtmlAttribute tableAttr in node.Attributes)
									{
									switch(tableAttr.Name)
										{
										//-get the **RowSpanning** if there is a value
										case "rowspan": //- get the rowspan if there is any
											{
											if(tableAttr.Value == null)
												tableRowSpanQty = 1;
											else
												{
												if(!int.TryParse(tableAttrValue, out tableRowSpanQty))
													tableRowSpanQty = 1;
												}
											break;
											}
										//-get the **ColumnSpanning** if there is a value
										case "colspan":
											{
											if(tableAttr.Value == null)
												tableColumnSpanQty = 1;
											else
												{
												if(!int.TryParse(tableAttrValue, out tableColumnSpanQty))
													tableColumnSpanQty = 1;
												}
											break;
											}
										//get the column width if it is specified
										case "style":
											{//-Get the column **Width** if it is specified
											if(tableAttr.Value.Contains("width"))
												{
												if(tableAttrValue.Contains("%"))
													{
													if(int.TryParse(tableAttrValue.Substring(
														tableAttrValue.IndexOf(":") + 2,
														(tableAttrValue.IndexOf("%") - tableAttrValue.IndexOf(":") - 2)),
														out columnWidth) == false)
														{
														//- Could not parse the integer which means the tableWidth remains as it was before the parse.
														}
													}
												else //-Column width is **NOT** a percentage :. px value
													{
													if(int.TryParse(tableAttrValue.Substring(
														tableAttrValue.IndexOf(":") + 2,
														(tableAttrValue.IndexOf("px") - tableAttrValue.IndexOf(":") - 2)),
														out tableWidth) == false)
														{
														//-Could not parse the integer which means the tableWidth remains as it was before the parse.
														}
													}
												}
											else
												{
												columnWidth = 0;
												}
											break;
											}
										}
									}
								}
							else //- the table Header Cell doesn't have any attributes
								{
								tableColumnSpanQty = 1;
								tableRowSpanQty = 1;
								columnWidth = 0;
								}
							Console.Write("\n\t\t<col> Header Column: Width: {0}\t RowSpan: {1}\t ColumnSpan: {2}\t", columnWidth, tableRowSpanQty, tableColumnSpanQty);
							break;
							}

						default:
							{
							//Console.WriteLine("\n ~~~ skip {0} ~~~", node.Name);
							continue;
							}
						}
					}

				Console.WriteLine("\n______________________________ HTML decoding iterations END ______________________________________");

				return true;
				}

			catch(InvalidTableFormatException exc)
				{
				Console.WriteLine("\n\nException: {0} - {1}", exc.Message, exc.Data);
				// Update the counters before returning
				parTableCaptionCounter = this.TableCaptionCounter;
				parImageCaptionCounter = this.ImageCaptionCounter;
				parPictureNo = this.PictureNo;
				parHyperlinkID = this.HyperlinkID;
				throw new InvalidContentFormatException(exc.Message);
				}

			catch(InvalidImageFormatException exc)
				{
				Console.WriteLine("\n\nException: {0} - {1}", exc.Message, exc.Data);
				// Update the counters before returning
				parTableCaptionCounter = this.TableCaptionCounter;
				parImageCaptionCounter = this.ImageCaptionCounter;
				parPictureNo = this.PictureNo;
				parHyperlinkID = this.HyperlinkID;
				throw new InvalidContentFormatException(exc.Message);
				}

			catch(Exception exc)
				{
				// Update the counters before returning
				parTableCaptionCounter = this.TableCaptionCounter;
				parImageCaptionCounter = this.ImageCaptionCounter;
				parPictureNo = this.PictureNo;
				parHyperlinkID = this.HyperlinkID;
				Console.WriteLine("\n**** Exception **** \n\t{0} - {1}\n\t{2}", exc.HResult, exc.Message, exc.StackTrace);
				throw new InvalidContentFormatException("An unexpected error occurred at this point, in the document generation. \nError detail: " + exc.Message);
				}
			finally
				{
				// Update the counters before returning
				parTableCaptionCounter = this.TableCaptionCounter;
				parImageCaptionCounter = this.ImageCaptionCounter;
				parPictureNo = this.PictureNo;
				parHyperlinkID = this.HyperlinkID;
				}
			}


		private static string CleanText(string parText, string parClientName)
			{
			//!The sequence in which these statements appear is important

			//-keep this code for debugging purposes if strange/unwanted characters appear.
			/*for(int i = 0; i < parText.Length; i++)
				{
				Console.Write("|" + ((int)parText[i]).ToString());
				}
			*/

			string cleanText = parText;
			cleanText = cleanText.Replace(((char)8203).ToString(), "");
			cleanText = cleanText.Replace("&#160;", " ");	//-remove *Hard space* characters
			cleanText = cleanText.Replace("<br>?", "");		//-Rich Text Break tags
			cleanText = cleanText.Replace("     ", " ");      //-cleanup any *5* spaces
			cleanText = cleanText.Replace("   ", " ");        //-cleanup any *triple* spaces
			cleanText = cleanText.Replace("  ", " ");         //-cleanup any *double* spaces
			cleanText = cleanText.Replace("\r", "");          //- remove carraige *return* characters
			cleanText = cleanText.Replace("\n", "");          //-remove *New Line* characters
			cleanText = cleanText.Replace("\t", "");          //-remove *Tab* characters
			if(cleanText == " " || cleanText == "  " || cleanText == "   ")
				cleanText = "";                              //-cleanup the string if it contains only a space.

			//-Replace ClientName #tag with actual value
			cleanText = cleanText.Replace("#ClientName#", parClientName);
			cleanText = cleanText.Replace("#clientname#", parClientName);
			cleanText = cleanText.Replace("#CLIENTNAME#", parClientName);

			return cleanText;
			}

		private static Tuple<int, int> GetBulletNumberLevels(string parXpath)
			{
			int bulletLevels = 0;
			int numberLevels = 0;
			int positionInString = 0;

			//+Check the number of **Headings** </h> tags
			//-Check if the xPath contains any bullets
			positionInString = 0;

			//+Check the number of **Unorganised List** </ul> tags
			//-Check if the xPath contains any bullets
			positionInString = 0;
			if(parXpath.Contains("/ul"))
				{ //- if it contains bullets, count the number of bullets
				for(int i = 0; i < parXpath.Length - 1;)
					{
					//-get the ocurrences of bullets
					positionInString = parXpath.IndexOf("/ul", i);
					if(positionInString >= 0)
						{
						bulletLevels += 1;
						i = positionInString + 3;
						}
					else if(positionInString < 0)
						break;
					}
				}
			else //-If it doesn't contain any bullets, set **bulletLevels** to 0
				{
				bulletLevels = 0;
				}

			//+Check the number of **Organised List** </ol> tags
			//-Check if the xPath contains any **Organised Lists** (Numbered tags).
			positionInString = 0;
			if(parXpath.Contains("/ol"))
				{ //- if it contains any, count the number of occurrences
				for(int i = 0; i < parXpath.Length - 1;)
					{
					//-get the ocurrences of tags
					positionInString = parXpath.IndexOf("/ol", i);
					if(positionInString >= 0)
						{
						numberLevels += 1;
						i = positionInString + 3;
						}
					else if(positionInString < 0)
						break;

					}
				}
			else //-If it doesn't contain any numbers, set **numberLevels** to 0
				{
				numberLevels = 0;
				}

			//-Return the counted occurrences
			return new Tuple<int, int>(bulletLevels, numberLevels);
			}

		//***R

		private void DetermineTableGrid(HtmlNodeCollection parHTMLnodes)
			{
			try
				{
				//- First clear the GridColumnWidths...
				this.TableToBeBuild.GridColumnWidths = new List<int>();

				//-Initialise Variables and Properties
				int columnCounter = 1;
				string colSpan = "";
				string columnWidthValue = "";
				int columnWidthPercentage = 0;
				int columnWidthPixels = 0;
				int columnRowSpanValue = 0;
				int columnSpan = 0;
				//- **1st** = Width Value, **2nd** = columnSpan value
				Tuple<int, int> columnWidth;
				List<Tuple<int,int>> tableGrid = new List<Tuple<int, int>>();

				
				//?This is what happens in this section of code...
				//- | 1. All the rows of the table is processed to determine the GRID of the table that need to be created in the MS Document.
				//- | 2. Each Column is each Row is inspected to determine how many columns is in the table and the **WIDTH** of each column.
				//-| 3. The complicating factors are:
				//- |     a) ... some columns may **NOT** have their width specified in the HTML.
				//-|     b) ...columns may **SPAN** multiple columns which means their width may not applu to each of the spanned columns..

				Console.WriteLine("\t\t\t {0} Begin to Define Table Grid {0}", new String('=', 40));
				//- Process the collection of columns that were send as parameter.
				foreach(HtmlNode node in parHTMLnodes)
					{
					switch(node.Name)
						{
						//+Table Body
						case "tbody":
							{
							columnCounter = 0;

							break;
							}

						//+Table Row
						case "tr":
							{
							columnCounter = 0;
							break;
							}

						//+Table Header
						case "th":
							{
							columnCounter += 1;
							colSpan = node.ChildAttributes("colspan").FirstOrDefault().Value;
							//-Check if the *cell* has a **colspan** value defined
							if(string.IsNullOrWhiteSpace(colSpan))
								{ //- The *cell* doesn't have a **colspan** value defined

								}
							else   //-The *cell* has a defined **colspan** value
								{
								//-Check if a column already Exist...
								if(tableGrid.Count < columnCounter)
									{
									

									}
								}

							//-Check if a cell width is defined...
							columnWidthValue = node.ChildAttributes("style").Where(v => v.Name == "width").FirstOrDefault().Value;
							//-Check if a **width** value was defined for the cell...
							if(String.IsNullOrWhiteSpace(columnWidthValue))
								{ //-The *cell* DOESn't have a defined **width**

								}
							else
								{ //- The *cell* HAS a defined **width**

								}

							//TODO: Assign the width and the colspan to the columnGrid
							tableGrid.Add();
							columnWidth = new Tuple<int, int>();

							break;
							}

						//+Table Cell
						case "td":
							{
							columnCounter += 1;

							break;
							}
						}


					if(node.Name == "tbody")
						{
						columnCounter = 1;
						}
					else 

					//+Determine the width of each column
					//-Check if a width is defined for the cell
					//-Does the *column* contains a **Width** attribute?
					if(node.getAttribute("width", 0) == null)
						{ //-it does not contain a *width* attribute...
						  //-Check if *width* is specified in a style
						if(node.style.width == null)
							{//- Width not specified in the style either...
							//!Report a Content Format Error!
							this.Table2insert.Active = false;
							throw new InvalidTableFormatException("The width of a tables's COLUMNS was NOT specified. Please review "
								+ "the table content and specify a table width % value.");
							}
						else //-A width value was found in the Style...
							{
							columnWidthValue = node.style.width;
							}
						}
					else //-Awidth value was found in the **Width attribute**
						{
						columnWidthValue = node.getAttribute("width", 0);
						}

					//-At this point there is a *column width* defined and the value resides in **sWidth**

					//-Determine HOW the table width is specified
					if(columnWidthValue.IndexOf("%", 0) > 0)
						{
						//this.TableColumnUnit = "%";
						columnWidthValue = columnWidthValue.Substring(0, columnWidthValue.IndexOf("%", 0));
						//-place the value into a temporary variable and remove any decimal if present...
						if(int.TryParse(columnWidthValue, out columnWidthPercentage))
							{
							columnWidthPixels = this.Table2insert.WidthPixels * (columnWidthPercentage / 100);
							}
						else //-If the value cannot be parsed, it means it is an invalid value...
							{
							//!Report a Content Error!
							this.Table2insert.Active = false;
							throw new InvalidTableFormatException("Column's width value is set as " + columnWidthPercentage
							+ "%, which is not correct. Table width values must not contain decimal values an preferably a percentage integer."
							+ " Please review the content and ensure all columns have a integer % width value.");
							}
						}
					else //- the width is specified in px...
						{ //-place the value into a temporary variable and remove any decimal if present...
						if(int.TryParse(columnWidthValue.Substring(0, (columnWidthValue.IndexOf(".", 0))), out columnWidthPixels))
							{//-the parse was successful..
							//-calulate the with percentage of the column.
							columnWidthPercentage = (columnWidthPixels / this.Table2insert.WidthPixels) * 100;
							}
						else //-The parse failed
							{
							this.Table2insert.Active = false;
							throw new InvalidTableFormatException("The table width value is set as " + columnWidthPercentage
							+ "px, which is outside the valid range of 1% to 100%. Please review the content and correct the "
							+ "table width % value.");
							}
						}
					//+Store the column width in the List of Column Percentages
					//-But before we save, let's check if there are any merged columns, to ensure the grid is complete

					if(node.getAttribute("colspan", 0) == null)
						{ //-There is no Column Merge present.
						columnSpan = 1;
						}
					else //-Extract the **colspan** value
						{
						columnRowSpanValue = node.getAttribute("colspan", 0);
						Console.WriteLine("ColumnSpan: {0}", columnRowSpanValue);
						//-Convert the string value to an integer
						//if(!int.TryParse(columnRowSpanValue,out columnSpan))
						//	{//-conversion failed..
						//	columnSpan = 1;
						//	}
						}
					
					//-Add a column each mearged column (**cell merges**) according to the value in the *columnSpan* variable
					for(int i = 0; i < columnSpan; i++)
						{
						this.Table2insert.GridColumnWidths.Add(columnWidthPixels);
						}
					}
				}
			catch(InvalidTableFormatException exc)
				{
				Console.WriteLine("Exception: {0} - {1}", exc.Message, exc.Data);
				this.Table2insert.Active = false;
				throw new InvalidTableFormatException(exc.Message);
				}

			catch(Exception exc)
				{
				Console.WriteLine("\n\nException ERROR: {0} - {1} - {2} - {3}", exc.HResult, exc.Source, exc.Message, exc.Data);
				}

			Console.WriteLine("\t\t\t {0} Table Grid Defined {0}", new String('=', 40));

			} // end of DetermineTableGrid
			*/
	
		//===G
		//++BuildCompleteTable

		public static DocumentFormat.OpenXml.Wordprocessing.Table BuildCompleteTable(
			WorkTable parWordTable,
			string parContentLayer)
			{
			TableRow objTableRow = new TableRow();
			TableCell objTableCell = new TableCell();
			Paragraph objParagraph = new Paragraph();
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();

			//+Create the OpenXML Table object instance...
			DocumentFormat.OpenXml.Wordprocessing.Table objTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
			//Construct the table in the 
			objTable = oxmlDocument.ConstructTable(
				parFirstRow: parWordTable.FirstRow,
				parFirstColumn: parWordTable.FirstColumn,
				parLastColumn: parWordTable.LastColumn,
				parLastRow: parWordTable.LastRow,
				parNoVerticalBand: true,
				parNoHorizontalBand: true);

			//+Construct the Table Grid
			TableGrid objTableGrid = new TableGrid();
			objTableGrid = oxmlDocument.ConstructTableGrid(parColumnWidthList: parWordTable.GridColumnWidths);
			objTable.Append(objTableGrid);

			//+ Process all the table Rows...
			foreach(WorkRow rowItem in parWordTable.Rows)
				{
				//-Get the Table's properties... 
				TableProperties objTableProperties = objTable.GetFirstChild<TableProperties>();
				//-Get the **TableLook** from in the TableProperties...
				TableLook objTableLook = objTableProperties.GetFirstChild<TableLook>();

				if(rowItem.FirstRoW)
					{
					//-Update the **FirsRow** value...
					objTableLook.FirstRow = true;
					}

				if(rowItem.LastRow)
					{
					//-Update the **LastRow** value...
					objTableLook.LastRow = true;
					}

				//- Create a **TableRow** object instance... row to the table
				objTableRow = new TableRow();
				objTableRow = oxmlDocument.ConstructTableRow(
					parIsFirstRow: rowItem.FirstRoW,
					parIsLastRow: rowItem.LastRow);

				//+Process all the Columns
				foreach(WorkColumn columnItem in rowItem.Columns)
					{
					//-Update the **TableLook** with First- and LastColumns id requried...
					if(columnItem.FirstColumn)
						objTableLook.FirstColumn = true;
					if(columnItem.LastColumn)
						objTableLook.LastColumn = true;

					//- Add the **TableColumn** to the row...
					objTableCell = new TableCell();
					objTableCell = oxmlDocument.ConstructTableCell(
						parCellWidth: columnItem.WidthPixel,
						parHasCondtionalFormatting: false,
						parIsFirstRow: rowItem.FirstRoW,
						parIsLastRow: rowItem.LastRow,
						parIsFirstColumn: columnItem.FirstColumn,
						parIsLastColumn: columnItem.LastColumn,
						parRowMerge: columnItem.SpanRows,
						parColumnMerge: columnItem.SpanColumns,
						parHorizontalAlignment: columnItem.AlignHorizontal.ToString(),
						parVerticalAlignment: columnItem.AlignmVertical.ToString());

					//- Add the Table Text
					objParagraph = new Paragraph();

					foreach(TextSegment objTextSegment in columnItem.ContentSegments)
						{
						objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
						objRun = oxmlDocument.Construct_RunText
								(parText2Write: objTextSegment.Text,
								parContentLayer: parContentLayer,
								parBold: objTextSegment.Bold,
								parItalic: objTextSegment.Italic,
								parUnderline: objTextSegment.Undeline,
								parSubscript: objTextSegment.Subscript,
								parSuperscript: objTextSegment.Superscript);

						objParagraph.Append(objRun);
						}
					objTableCell.Append(objParagraph);

					objTableRow.Append(objTableCell);
					}
				//-once all the **TableColumns are added, append the the TableRow to the Table object.
				objTable.Append(objTableRow);
				}

			return objTable;
			}

		/*

		public static String CleanHTMLstring(
			string parHTML2Decode, string parClientName)
			{
			string cleanText = String.Empty;
			string WorkString = String.Empty;
			IHTMLDocument2 objHTMLDocument2 = (IHTMLDocument2)new HTMLDocument();
			objHTMLDocument2.write(parHTML2Decode);
			IHTMLElementCollection objElementCollection = objHTMLDocument2.body.children;

			foreach(IHTMLElement objHTMLelement in objElementCollection)
				{
				WorkString = CleanString(objHTMLelement.innerText, parClientName);
				if(WorkString != String.Empty)
					{
					cleanText += WorkString;
					}
				}

			// replace and/or remove/replace unwanter characthers from the string...
			cleanText = HTMLdecoder.CleanString(cleanText, parClientName);

			return cleanText;
			} //end CleanHTMLstring method

	*/
		public static String CleanString(
			string parTextToClean,
			string parClientName)
			{
			//- replace and/or remove special strings from text...
			if(parTextToClean == null
			|| parTextToClean == String.Empty)
				{
				parTextToClean = String.Empty;
				}
			else
				{
				parTextToClean = parTextToClean.Replace(((char)8203).ToString(), "");
				parTextToClean = parTextToClean.Replace("&quot;", Convert.ToString(value: (char)22));
				parTextToClean = parTextToClean.Replace("&nbsp;", " ");     //- replace hard spaces
				parTextToClean = parTextToClean.Replace("&#160;", "");      //- Remove HTML hard space
				parTextToClean = parTextToClean.Replace("&#39;", "'");      //- Replace HTML single quote with a '
				parTextToClean = parTextToClean.Replace("     ", " ");      //- replace 5 consecutive spaces with 1 space
				parTextToClean = parTextToClean.Replace("    ", " ");       //- replace 4 consecutive spaces with 1 space
				parTextToClean = parTextToClean.Replace("   ", " ");        //- replace 3 consecutive spaces with 1 space
				parTextToClean = parTextToClean.Replace("  ", " ");         //- replace 2 consecutive spaces with 1 space
				parTextToClean = parTextToClean.Replace("\n", "");
				parTextToClean = parTextToClean.Replace("\r", "");
				parTextToClean = parTextToClean.Replace("\t", "");
				if(parTextToClean == " " || parTextToClean == "  " || parTextToClean == "   ")
					parTextToClean = "";

				if(parClientName != "")
					{
					parTextToClean = parTextToClean.Replace("#ClientName#", parClientName);
					parTextToClean = parTextToClean.Replace("#clientname#", parClientName);
					parTextToClean = parTextToClean.Replace("#CLIENTNAME#", parClientName);
					}
				}
			return parTextToClean;

			}    //- end of Method



		}


	//***
	//++ These classes are used to construct tables
	public class WorkTable
		{
		public enum enumWidthType
			{
			Pixel,
			Percentage
			}

		/// <summary>
		/// This flag MUST be set to TRUE as soon as a WorkTable object is instanciated.
		/// </summary>
		public bool Active { get; set; } = false;

		/// <summary>
		/// May contain a string which can be used as the Table's Caption. 
		/// </summary>
		public String Caption { get; set; } = String.Empty;

		/// <summary>
		/// Needs to be set to indicate the WidthType of the the original HTML table. Will be used during the calculation of the Table's width scaling...
		/// </summary>
		public enumWidthType OriginalTableWidthType { get; set; } = enumWidthType.Percentage;

		/// <summary>
		/// The original value of the table width as it was defined in the HTML source...
		/// </summary>
		public int OriginalTableWidthValue { get; set; } = 0;

		/// <summary>
		/// The percentage of the Document page that the table will occupy...
		/// </summary>
		public int WidthPrecetage { get; set; } = 0;


		public int WidthPixels { get; set; } = 0;

		/// <summary>
		/// Set to TRUE (the default is FALSE) as soon as the Table Grid was determined and defined.
		/// </summary>
		public bool GridDone { get; set; } = false;

		/// <summary>
		/// Set to True (the default is FALSE) if the table has a FIRST or HEADER Row defined.
		/// </summary>
		public bool FirstRow { get; set; } = false;

		/// <summary>
		/// Set to TRUE (the default is FALSE) if the table has a Last or FOOTER Row defined.
		/// </summary>
		public bool LastRow { get; set; } = false;

		/// <summary>
		/// Set to TRUE (the default is FALSE) if the table has a FIRST Column defined.
		/// </summary>
		public bool FirstColumn { get; set; } = false;

		/// <summary>
		/// Set to TRUE (the default is FALSE) if the table has a LAST or TOTAL column defined.
		/// </summary>
		public bool LastColumn { get; set; } = false;

		/// <summary>
		/// This list of integers defined the table's column grids from the left to the right column...
		/// </summary>
		public List<int> GridColumnWidths { get; set; } = new List<int>();

		/// <summary>
		/// This list is a collection of WorkRow objects which represents each row of the table from top to bottom of the table...
		/// </summary>
		public List<WorkRow> Rows { get; set; } = new List<WorkRow>();

		}


	//++
	public class WorkRow
		{
		public int Number { get; set; } = 0;
		public bool FirstRoW { get; set; } = false;
		public bool OddRow { get; set; } = false;
		public bool EvenRow { get; set; } = false;
		public bool LastRow { get; set; } = false;
		public List<WorkColumn> Columns { get; set; } = new List<WorkColumn>();
		}

	public class WorkColumn
		{
		public enum enumAlignmentHorizontal
			{
			Left = 1,
			Centre = 2,
			Right = 3
			}

		public enum enumAlignmentVertical
			{
			Top = 1,
			Centre = 2,
			Bottom = 3
			}

		public int Number { get; set; } = 0;
		public bool FirstColumn { get; set; } = false;
		public bool LastColumn { get; set; } = false;
		public int WidthPercentage { get; set; } = 0;
		public int WidthPixel { get; set; } = 0;
		public int SpanColumns { get; set; } = 0;
		public int SpanRows { get; set; } = 0;
		public enumAlignmentHorizontal AlignHorizontal { get; set; } = enumAlignmentHorizontal.Left;
		public enumAlignmentVertical AlignmVertical { get; set; } = enumAlignmentVertical.Centre;
		public bool Bold { get; set; } = false;
		public bool Italic { get; set; } = false;
		public bool Underline { get; set; } = false;
		public bool LineThrough { get; set; } = false;
		public bool Subscript { get; set; } = false;
		public bool Superscript { get; set; } = false;
		public string HTMLcontent { get; set; } = String.Empty;
		public List<TextSegment> ContentSegments { get; set; } = new List<TextSegment>();

		}

	}
