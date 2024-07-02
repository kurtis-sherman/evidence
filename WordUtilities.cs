using System;
using System.Configuration;
using System.Drawing.Imaging;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace evidence
{
    public class WordUtilities
    {
        public enum AnnotationType
        {
            None,
            Info,
            Pass,
            Fail
        }

        private string docFolderPath;
        static string docFilename;

        public WordUtilities()
        {
            // Initialize or read docFolderPath from app.config
            docFolderPath = ConfigurationManager.AppSettings["DocFolderPath"];
            // Use default path if not specified
            if (string.IsNullOrWhiteSpace(docFolderPath))
            {
                string currentUser = Environment.UserName;
                string currentUserDocuments = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), currentUser);
                docFolderPath = Path.Combine(currentUserDocuments, "Documents"); // Default folder path
            }

            // Ensure docFolderPath exists, create if it doesn't
            EnsureFolderExists(docFolderPath);
        }

        public string CreateWordDocument()
        {
            // Generate timestamp for the Word document
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            // Sanitize filename
            string sanitizedName = SanitizeFilename($"evidence_{timestamp}");

            // Create Word document
            docFilename = Path.Combine(docFolderPath, $"{sanitizedName}.docx");

            // Initialize Word application and document
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();

            try
            {
                // Set page layout to landscape
                wordDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;

                // Set margins to 0.25 inches (18 points)
                float marginInches = 0.25f;
                float marginPoints = marginInches * 72; // 1 inch = 72 points
                wordDoc.PageSetup.LeftMargin = marginPoints;
                wordDoc.PageSetup.RightMargin = marginPoints;
                wordDoc.PageSetup.TopMargin = marginPoints;
                wordDoc.PageSetup.BottomMargin = marginPoints;

                // Save the document without displaying any dialogs
                wordDoc.SaveAs2(docFilename, Word.WdSaveFormat.wdFormatDocumentDefault,
                                AddToRecentFiles: false, Password: "", WritePassword: "",
                                ReadOnlyRecommended: false, EmbedTrueTypeFonts: false,
                                SaveNativePictureFormat: false, SaveFormsData: false,
                                SaveAsAOCELetter: false, Encoding: Type.Missing,
                                InsertLineBreaks: Type.Missing, AllowSubstitutions: false,
                                LineEnding: Type.Missing, AddBiDiMarks: false,
                                CompatibilityMode: Word.WdCompatibilityMode.wdCurrent
                                );
                return docFilename;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating Word document: {ex.Message}");
                return null; // Handle error gracefully in your application
            }
            finally
            {
                // Close and quit Word application
                wordDoc.Close();
                wordApp.Quit();

                // Release COM objects to avoid memory leaks
                ReleaseObject(wordDoc);
                ReleaseObject(wordApp);
            }
        }


        private string SanitizeFilename(string filename)
        {
            // Replace invalid filename characters with underscores
            return Regex.Replace(filename, "[\\\\/:*?\"<>|]", "_");
        }

        private void EnsureFolderExists(string folderPath)
        {
            // Create the folder if it doesn't exist
            if (!Directory.Exists(folderPath))
            {
                try
                {
                    Directory.CreateDirectory(folderPath);
                    Console.WriteLine($"Created folder: {folderPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error creating folder {folderPath}: {ex.Message}");
                    // Handle error gracefully in your application
                }
            }
        }

        public Bitmap CaptureScreen()
        {
            foreach (Screen screen in Screen.AllScreens)
            {
                if (screen.Bounds.Contains(Cursor.Position))
                {
                    Bitmap screenshot = new Bitmap(screen.Bounds.Width, screen.Bounds.Height, PixelFormat.Format32bppArgb);
                    using (Graphics g = Graphics.FromImage(screenshot))
                    {
                        g.CopyFromScreen(screen.Bounds.X, screen.Bounds.Y, 0, 0, screen.Bounds.Size, CopyPixelOperation.SourceCopy);
                    }
                    return screenshot;
                }
            }
            return null;
        }

        private Bitmap AnnotateScreenshot(Bitmap screenshot, string annotationText, Point annotationPosition, AnnotationType type)
        {
            Bitmap annotated = new Bitmap(screenshot);

            using (Graphics g = Graphics.FromImage(annotated))
            {
                int rectSize = 40;
                int rectOffset = rectSize / 2;
                Screen scnInContext = Screen.FromPoint(annotationPosition);

                int scnX = annotationPosition.X - scnInContext.Bounds.X;
                int scnY = annotationPosition.Y - scnInContext.Bounds.Y;

                int rectX = Math.Max(0, scnX - rectOffset);
                int rectY = Math.Max(0, scnY - rectOffset);

                Rectangle rect = new Rectangle(rectX, rectY, rectSize, rectSize);
                Color colTransp = Color.FromArgb(128, 0, 0, 255);

                if (rect.X >= 0 && rect.Y >= 0 && rect.X + rect.Width <= annotated.Width && rect.Y + rect.Height <= annotated.Height)
                {
                    g.FillRectangle(new SolidBrush(colTransp), rect);
                }
                else
                {
                    throw new Exception("Rectangle out of bounds");
                }

                //Font wingdingsFont = new Font("Wingdings", 20); // Adjust font size as needed
                //char checkboxSymbol = (char)0xFC;  // Unicode character code for checkbox in Wingdings font

                //Font wingdingsFont = new Font("Wingdings", 20); // Adjust font size as needed
                //char checkboxSymbol = (char)0xFC;  // Unicode character code for checkbox in Wingdings font

                // Draw checkbox symbol at annotationPosition
                //Point textPosition = new Point(annotationPosition.X + 20 , annotationPosition.Y);
                //g.DrawString(checkboxSymbol.ToString(), wingdingsFont, Brushes.Green, annotationPosition);

                Font fontD = SystemFonts.DefaultFont;
                char symbol = ' ';
                Brush brush = Brushes.White;
                switch (type)
                {
                    case AnnotationType.Info:
                        fontD = new Font("Webdings", 20);
                        symbol = 'i';
                        break;
                    case AnnotationType.Pass:
                        fontD = new Font("Wingdings 2", 20);
                        symbol = 'R';
                        brush = Brushes.Green;
                        break;
                    case AnnotationType.Fail:
                        fontD = new Font("Wingdings 2", 20);
                        symbol = 'Q';
                        brush = Brushes.Red;
                        break;
                }

                // Load values from app.config with defaults
                int annoSymOffsetX = int.TryParse(ConfigurationManager.AppSettings["AnnotationSymbolOffsetX"], out int offsetX) ? offsetX : 20;
                int annoSymOffsetY = int.TryParse(ConfigurationManager.AppSettings["AnnotationSymbolOffsetY"], out int offsetY) ? offsetY : 20;

                // Draw text at annotationPosition with offsets
                Point textPosition = new Point(annotationPosition.X + annoSymOffsetX, annotationPosition.Y + annoSymOffsetY);

                // Use textPosition in your DrawString method
                g.DrawString(symbol.ToString(), fontD, brush, textPosition);
                //g.DrawString(annotationText, SystemFonts.DefaultFont, Brushes.Red, annotationPosition);


                //g.DrawString(annotationText, SystemFonts.DefaultFont, Brushes.Red, annotationPosition);
            }

            return annotated;
        }

        public void AppendScreenshotToWord(AnnotationType type, string annotationText)
        {
            // Initialize Word application
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Open(docFilename);

            // Capture screenshot and annotate if annotationText is not empty
            Bitmap screenshot = CaptureScreen();
            Bitmap annotated = AnnotateScreenshot(screenshot, annotationText, Cursor.Position, type);

            // Save annotated image to a temporary file
            string tempImagePath = Path.GetTempFileName();
            annotated.Save(tempImagePath, ImageFormat.Png);

            // Move to end of document
            object missing = System.Reflection.Missing.Value;
            Word.Range endRange = wordDoc.Content;
            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            endRange.InsertParagraphAfter();

            // Insert annotationText as Arial font size 14
            Word.Paragraph textPara = wordDoc.Content.Paragraphs.Add();
            textPara.Range.Text = annotationText;
            textPara.Range.Font.Name = "Arial";
            textPara.Range.Font.Size = 14;
            textPara.Range.InsertParagraphAfter();

            // Insert horizontal line
            Word.Paragraph horizLine = wordDoc.Content.Paragraphs.Add();
            horizLine.Range.InlineShapes.AddHorizontalLineStandard();
            horizLine.Range.InsertParagraphAfter();

            // Insert image into Word document
            Word.InlineShape inlineShape = horizLine.Range.InlineShapes.AddPicture(tempImagePath);

            // Calculate maximum width that fits within page margins
            float maxWidth = wordDoc.PageSetup.PageWidth - wordDoc.PageSetup.LeftMargin - wordDoc.PageSetup.RightMargin;

            // Resize image if necessary
            if (inlineShape.Width > maxWidth)
            {
                float scaleRatio = maxWidth / inlineShape.Width;
                inlineShape.Width *= scaleRatio;
                inlineShape.Height *= scaleRatio;
            }

            // Clean up
            File.Delete(tempImagePath);

            // Save and close the Word document
            wordDoc.Save();
            wordDoc.Close();
            wordApp.Quit();
        }



        //public void AppendScreenshotToWord(AnnotationType type, string annotationText)
        //{
        //    Bitmap screenshot = CaptureScreen(); // Capture screenshot
        //    Point currentMousePosition = Cursor.Position; // Get current mouse position

        //    // Initialize Word application
        //    Word.Application wordApp = new Word.Application();
        //    Word.Document wordDoc = wordApp.Documents.Open(docFilename);

        //    // Annotate screenshot if annotationText is not empty
        //    Bitmap annotated = AnnotateScreenshot(screenshot, annotationText, currentMousePosition, type);

        //    // Save annotated image to a temporary file
        //    string tempImagePath = Path.GetTempFileName();
        //    annotated.Save(tempImagePath, ImageFormat.Png);

        //    // Move to end of document
        //    object missing = System.Reflection.Missing.Value;
        //    wordApp.Selection.EndKey(Word.WdUnits.wdStory, missing);

        //    // Insert image into Word document
        //    Word.InlineShape inlineShape = wordDoc.InlineShapes.AddPicture(tempImagePath);
        //    inlineShape.Width = screenshot.Width;
        //    inlineShape.Height = screenshot.Height;
        //    inlineShape.Range.Cut();

        //    // Horizontal line
        //    Word.Paragraph horizLine = wordDoc.Content.Paragraphs.Add();
        //    horizLine.Range.InlineShapes.AddHorizontalLineStandard();

        //    Word.Selection sel = wordApp.Selection;

        //    sel.Paste();

        //    // Clean up temporary file
        //    File.Delete(tempImagePath);

        //    // Save and close the Word document
        //    wordDoc.Save();
        //    wordDoc.Close();
        //    wordApp.Quit();
        //}

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine($"Error releasing object: {ex.Message}");
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
