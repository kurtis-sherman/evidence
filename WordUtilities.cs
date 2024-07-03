using System;
using System.Configuration;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace evidence
{
    public class WordUtilities
    {
        private Word.Application wordApp;
        private Word.Document wordDoc;
        private string docFolderPath;
        private static string docFilename;

        public enum AnnotationType
        {
            None,
            Info,
            Pass,
            Fail
        }

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

            // Generate timestamp for the Word document
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            // Sanitize filename
            string sanitizedName = SanitizeFilename($"evidence_{timestamp}");

            // Create Word document
            docFilename = Path.Combine(docFolderPath, $"{sanitizedName}.docx");

            // Initialize Word application and document
            InitializeWord();
        }

        private void InitializeWord()
        {
            // Check if Word application and document need initialization
            if (wordApp == null || wordDoc == null)
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Add();

                // Set page layout to landscape
                wordDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;

                // Set margins to 0.25 inches (18 points)
                float marginInches = 0.25f;
                float marginPoints = marginInches * 72; // 1 inch = 72 points
                wordDoc.PageSetup.LeftMargin = marginPoints;
                wordDoc.PageSetup.RightMargin = marginPoints;
                wordDoc.PageSetup.TopMargin = marginPoints;
                wordDoc.PageSetup.BottomMargin = marginPoints;
            }
        }

        public void AppendScreenshotToWord(AnnotationType type, string annotationText)
        {
            try
            {
                // Capture screenshot and annotate
                Bitmap screenshot = CaptureScreen();
                if (screenshot != null)
                {
                    // Determine the screen where the cursor was when the screenshot was taken
                    Screen cursorScreen = Screen.FromPoint(Cursor.Position);

                    // Adjust cursor position relative to the top-left corner of the monitor
                    int cursorXRelativeToMonitor = Cursor.Position.X - cursorScreen.Bounds.X;
                    int cursorYRelativeToMonitor = Cursor.Position.Y - cursorScreen.Bounds.Y;

                    // Get cursor position relative to the monitor where the mouse was when the screenshot was taken
                    Point cursorPositionRelativeToMonitor = new Point(cursorXRelativeToMonitor, cursorYRelativeToMonitor);

                    // Annotate the screenshot based on annotationText and annotation type
                    Bitmap annotated = AnnotateScreenshot(screenshot, annotationText, cursorPositionRelativeToMonitor, type);

                    // Save annotated image to a temporary file
                    string tempImagePath = Path.GetTempFileName();
                    annotated.Save(tempImagePath, ImageFormat.Png);

                    // Insert annotationText into Word document
                    AddTextToWordDocument(annotationText);

                    // Insert horizontal line
                    AddHorizontalLineToWordDocument();

                    // Insert image into Word document
                    AddImageToWordDocument(tempImagePath);

                    // Clean up temporary image file
                    File.Delete(tempImagePath);
                }
                else
                {
                    Console.WriteLine("Error: Unable to capture screenshot.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error appending screenshot to Word document: {ex.Message}");
                // Handle the exception as needed in your application
            }
        }

        public void AddTextToWordDocument(string text)
        {
            try
            {
                // Check if Word application and document are initialized
                if (wordDoc != null && wordApp != null)
                {
                    // Move to the end of the document
                    Word.Range endRange = wordDoc.Content;
                    endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    // Insert text into the document
                    Word.Paragraph para = wordDoc.Content.Paragraphs.Add();
                    para.Range.Text = text;
                    para.Range.InsertParagraphAfter();

                    // Optionally, you can format the text
                    para.Range.Font.Name = "Arial";
                    para.Range.Font.Size = 12;
                }
                else
                {
                    throw new InvalidOperationException("Word application or document is not initialized.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding text to Word document: {ex.Message}");
                // Handle the exception as needed in your application
            }
        }

        public void AddHorizontalLineToWordDocument()
        {
            try
            {
                // Check if Word application and document are initialized
                if (wordDoc != null && wordApp != null)
                {
                    // Move to the end of the document
                    Word.Range endRange = wordDoc.Content;
                    endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    // Insert horizontal line
                    Word.Paragraph horizLine = wordDoc.Content.Paragraphs.Add();
                    horizLine.Range.InlineShapes.AddHorizontalLineStandard();
                    horizLine.Range.InsertParagraphAfter();
                }
                else
                {
                    throw new InvalidOperationException("Word application or document is not initialized.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding horizontal line to Word document: {ex.Message}");
                // Handle the exception as needed in your application
            }
        }

        private void AddImageToWordDocument(string imagePath)
        {
            try
            {
                // Check if Word application and document are initialized
                if (wordDoc != null && wordApp != null)
                {
                    // Move to the end of the document
                    Word.Range endRange = wordDoc.Content;
                    endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    endRange.InsertParagraphAfter();

                    // Insert image into Word document
                    Word.InlineShape inlineShape = endRange.InlineShapes.AddPicture(imagePath, System.Reflection.Missing.Value, true, System.Reflection.Missing.Value);
                    //endRange.InlineShapes.AddPicture(imagePath, System.Reflection.Missing.Value, true, System.Reflection.Missing.Value);

                    // Calculate maximum width that fits within page margins
                    float maxWidth = wordDoc.PageSetup.PageWidth - wordDoc.PageSetup.LeftMargin - wordDoc.PageSetup.RightMargin;

                    // Resize image if necessary
                    if (inlineShape.Width > maxWidth)
                    {
                        float scaleRatio = maxWidth / inlineShape.Width;
                        inlineShape.Width *= scaleRatio;
                        inlineShape.Height *= scaleRatio;
                    }
                }
                else
                {
                    throw new InvalidOperationException("Word application or document is not initialized.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding image to Word document: {ex.Message}");
                // Handle the exception as needed in your application
            }
        }

        private Bitmap CaptureScreen()
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
                // Read circle radius from app.config with default value
                int circleRadius = int.TryParse(ConfigurationManager.AppSettings["CircleRadius"], out int radius) ? radius : 20;
                int circleDiameter = circleRadius * 2;

                // Calculate circle position centered at annotationPosition
                int circleX = annotationPosition.X - circleRadius;
                int circleY = annotationPosition.Y - circleRadius;

                // Create circle bounds
                Rectangle circleBounds = new Rectangle(circleX, circleY, circleDiameter, circleDiameter);

                // Calculate center of the circle
                int centerX = annotationPosition.X;
                int centerY = annotationPosition.Y;

                // Add a red pixel at the center of the circle
                if (centerX >= 0 && centerX < annotated.Width && centerY >= 0 && centerY < annotated.Height)
                {
                    annotated.SetPixel(centerX, centerY, Color.Red);
                }

                // Select font and symbol based on AnnotationType (configurable via app.config)
                Font font;
                char symbol;
                Brush brush = Brushes.White; // Default brush

                switch (type)
                {
                    case AnnotationType.Info:
                        font = new Font(ConfigurationManager.AppSettings["InfoFontName"] ?? "Webdings", 20);
                        string infoSymbolConfig = ConfigurationManager.AppSettings["InfoSymbol"];
                        symbol = !string.IsNullOrEmpty(infoSymbolConfig) ? infoSymbolConfig[0] : 'i'; // Default symbol 'i'
                        brush = new SolidBrush(ColorTranslator.FromHtml(ConfigurationManager.AppSettings["InfoColor"] ?? "#FFFFFF")); // Default: White
                        break;
                    case AnnotationType.Pass:
                        font = new Font(ConfigurationManager.AppSettings["PassFontName"] ?? "Wingdings 2", 20);
                        string passSymbolConfig = ConfigurationManager.AppSettings["PassSymbol"];
                        symbol = !string.IsNullOrEmpty(passSymbolConfig) ? passSymbolConfig[0] : 'R'; // Default symbol 'R'
                        brush = new SolidBrush(ColorTranslator.FromHtml(ConfigurationManager.AppSettings["PassColor"] ?? "#00FF00")); // Default: Green
                        break;
                    case AnnotationType.Fail:
                        font = new Font(ConfigurationManager.AppSettings["FailFontName"] ?? "Wingdings 2", 20);
                        string failSymbolConfig = ConfigurationManager.AppSettings["FailSymbol"];
                        symbol = !string.IsNullOrEmpty(failSymbolConfig) ? failSymbolConfig[0] : 'Q'; // Default symbol 'Q'
                        brush = new SolidBrush(ColorTranslator.FromHtml(ConfigurationManager.AppSettings["FailColor"] ?? "#FF0000")); // Default: Red
                        break;
                    default:
                        font = SystemFonts.DefaultFont;
                        symbol = ' ';
                        break;
                }

                // Read transparency level from app.config with default value
                int transparencyLevel = int.TryParse(ConfigurationManager.AppSettings["CircleTransparency"], out int transparency) ? transparency : 128;

                // Fill circle with transparent color
                Color transparentColor = Color.FromArgb(transparencyLevel, ((SolidBrush)brush).Color);
                g.FillEllipse(new SolidBrush(transparentColor), circleBounds);

                // Read symbol offsets from app.config with default values
                int symbolXOffset = int.TryParse(ConfigurationManager.AppSettings["SymbolXOffset"], out int xOffset) ? xOffset : 0;
                int symbolYOffset = int.TryParse(ConfigurationManager.AppSettings["SymbolYOffset"], out int yOffset) ? yOffset : 0;

                // Adjust symbol offsets if they push the symbol outside the annotated area
                int symbolX = annotationPosition.X + symbolXOffset; // Center horizontally
                int symbolY = annotationPosition.Y + symbolYOffset; // Center vertically

                // Ensure symbol is within bounds of the annotated area
                if (symbolX < 0)
                {
                    symbolX = 0;
                }
                else if (symbolX > annotated.Width - font.Height) // Adjust based on symbol size (assuming square font)
                {
                    symbolX = annotated.Width - font.Height;
                }

                if (symbolY < 0)
                {
                    symbolY = 0;
                }
                else if (symbolY > annotated.Height - font.Height) // Adjust based on symbol size (assuming square font)
                {
                    symbolY = annotated.Height - font.Height;
                }

                Point symbolPosition = new Point(symbolX, symbolY);
                g.DrawString(symbol.ToString(), font, brush, symbolPosition);
            }

            return annotated;
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

        public void SaveAndCloseWordDocument()
        {
            if (wordDoc != null && wordApp != null)
            {
                try
                {
                    // Save the document without displaying any dialogs
                    wordDoc.SaveAs2(docFilename, Word.WdSaveFormat.wdFormatDocumentDefault,
                                    AddToRecentFiles: false, Password: "", WritePassword: "",
                                    ReadOnlyRecommended: false, EmbedTrueTypeFonts: false,
                                    SaveNativePictureFormat: false, SaveFormsData: false,
                                    SaveAsAOCELetter: false, Encoding: Type.Missing,
                                    InsertLineBreaks: Type.Missing, AllowSubstitutions: false,
                                    LineEnding: Type.Missing, AddBiDiMarks: false,
                                    CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving Word document: {ex.Message}");
                }
                finally
                {
                    // Close and quit Word application
                    CloseWordDocument();
                }
            }
        }

        private void CloseWordDocument()
        {
            try
            {
                // Close and quit Word application
                wordDoc.Close(SaveChanges: true);
                wordApp.Quit(SaveChanges: true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error closing Word document: {ex.Message}");
            }
            finally
            {
                // Release COM objects to avoid memory leaks
                ReleaseObject(wordDoc);
                ReleaseObject(wordApp);
            }
        }

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
