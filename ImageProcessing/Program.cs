using Emgu.CV.CvEnum;
using Emgu.CV;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Emgu.CV.Structure;
using System.Drawing;
using Emgu.CV.Util;
using static System.Net.Mime.MediaTypeNames;
using Emgu.CV.Ocl;
using Emgu.CV.Cuda;
using System.IO;
using Emgu.CV.OCR;
using Tesseract;
namespace ImageProcessingPart1
{
    internal class Program
    {
        private static string? ImagePath = @"D:\Channel\Code Sell\Working\Input";
        static void Main(string[] args)
        {
            PreProcessingImage.Instance.ProcessImage(ImagePath);
        }
    }

    public class PreProcessingImage
    {
        private string? ImagePath = string.Empty;
        private static PreProcessingImage? instance = null;
        private Mat? OrignalImage;
        private Mat? result;
        private double BinaryThreshold = 150;
        private double MaxValue = 255;
        private string? datapath = @"D:\Channel\Code Sell\Working\Git\ImageProcessing\References\Tesseract";
        private static string? OutputPath = @"D:\Channel\Code Sell\Working\Output\test.png";
        StringBuilder stringBuilder = new StringBuilder();
        public static PreProcessingImage Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new PreProcessingImage();
                }
                return instance;
            }
        }

        public void ProcessImage(string imgPath)
        {
            DirectoryInfo? directoryInfo = new DirectoryInfo(imgPath);
            foreach (FileInfo fileInfo in directoryInfo.GetFiles())
            {
                ImagePath = fileInfo.FullName;
                ProcessImage(ImagePath, new Rectangle());
            }
        }
        public void ProcessImage(string filename, Rectangle rectangle)
        {
            ImagePath = filename;
            //Step 1: Read the image 1
            result = GetImage(ImreadModes.Grayscale);

            result = ResizeImage(result, result, 3000, 3000);

            // Step 2: Remove Non ZeroArea 2
            result = RemoveNonZeroAreafromImage(result);

            // Step 3: Convert to Binary 1
            result = GrayscaleToBinaryThreshold(result, BinaryThreshold, MaxValue, false);
            OrignalImage = result;
            result = GetLines(OrignalImage);

            // Step 4: Set Border 2
            result = SetImageBorder(result);

            // Step 5: Horizontal Vertical 3
            result = GetHorizontalVertical(result, result.Rows / 150);

            OrignalImage = result;
            OrignalImage.Save(OutputPath);

            // Step 6: OCR (Image to Text)
            foreach (FileInfo fileInfo in new DirectoryInfo(@"D:\Channel\Code Sell\Working\Output").GetFiles())
            {               
                ConvertImageToTextOCR(fileInfo.FullName);
            }
            File.WriteAllText(@"D:\Channel\Code Sell\Working\Output\OCRTEXT.txt", stringBuilder.ToString());

        }
        private Mat GetImage(ImreadModes imreadModes = ImreadModes.Unchanged)
        {
            Mat? Imagesize = new Mat(ImagePath);
            if (Imagesize.NumberOfChannels == 4 || Imagesize.NumberOfChannels == 3)
            {
                CvInvoke.CvtColor(Imagesize, Imagesize, ColorConversion.Bgr2Gray);
            }
            else
            {
                CvInvoke.CvtColor(Imagesize, Imagesize, ColorConversion.Gray2Bgr);
            }
            return Imagesize;
        }
        private Rectangle GetNonWhiteBounds(Bitmap bmp)
        {
            int minX = bmp.Width, minY = bmp.Height, maxX = 0, maxY = 0;
            Color white = Color.FromArgb(255, 255, 255);

            for (int y = 0; y < bmp.Height; y++)
            {
                for (int x = 0; x < bmp.Width; x++)
                {
                    Color pixel = bmp.GetPixel(x, y);
                    if (pixel.ToArgb() != white.ToArgb()) // Change this for tolerance
                    {
                        if (x < minX) minX = x;
                        if (y < minY) minY = y;
                        if (x > maxX) maxX = x;
                        if (y > maxY) maxY = y;
                    }
                }
            }

            if (minX > maxX || minY > maxY)
            {
                return new Rectangle(0, 0, 1, 1);
            }
            return new Rectangle(minX, minY, maxX - minX + 0, maxY - minY + 0);
        }
        private Mat RemoveNonZeroAreafromImage(Mat ThresholdImage)
        {
            using (Bitmap original = ThresholdImage.ToBitmap())
            {
                Rectangle cropRect = GetNonWhiteBounds(original);
                using (Bitmap cropped = original.Clone(cropRect, original.PixelFormat))
                {
                    return cropped.ToMat();
                }
            }
        }
        private Mat SetImageBorder(Mat ThresholdImage)
        {
            Image<Bgr, byte>? image = ThresholdImage.ToImage<Bgr, byte>();
            VectorOfPoint? points = new VectorOfPoint();
            CvInvoke.FindNonZero(image[ThresholdImage.NumberOfChannels].Mat, points);
            var lBoundingRectangle = new Mat(image[ThresholdImage.NumberOfChannels].Mat, CvInvoke.BoundingRectangle(points));

            int startRow = ((image.Rows - lBoundingRectangle.Rows) / 2) - 0;
            Rectangle rectangle = new Rectangle(1, startRow, image.Cols + 10, lBoundingRectangle.Rows + 10);
            image.Draw(rectangle, new Bgr(Color.White), 50,LineType.EightConnected);
            return image.Mat;
        }
        private Mat ResizeImage(Mat SourceImage, Mat SourceDestination, int Width, int Height)
        {
            CvInvoke.Resize(SourceImage, SourceDestination, new Size(Width, Height), 2, 2, Inter.Cubic);
            return SourceDestination;
        }
        private Mat GrayscaleToBinaryThreshold(Mat ThresholdImage, double binaryThreshold = 255, double Maxvalue = 255, bool IsCounter = false)
        {
            Image<Gray, byte>? toImage = ThresholdImage.ToImage<Gray, byte>();
            if (!IsCounter)
                CvInvoke.Threshold(toImage, toImage, binaryThreshold, Maxvalue, ThresholdType.BinaryInv);
            else

            CvInvoke.AdaptiveThreshold(toImage, toImage, Maxvalue, AdaptiveThresholdType.GaussianC, ThresholdType.BinaryInv, 255, 0);
            if (!IsCounter)
            {
                toImage.Erode(1);
                toImage.Dilate(1);
                toImage.PyrDown();
                toImage.PyrUp();
                toImage.Canny(2, 6);
                toImage = toImage.SmoothGaussian(1, 1, 5, 5);//1, 1, 5, 5
                toImage = toImage.SmoothBilateral(5, 5, 5);
                toImage = toImage.SmoothBlur(1, 1, false);
            }
            return toImage.Mat;
        }
  
        private RotateFlipType OrientationToFlipType(int orientation)
        {
            return orientation switch
            {
                6 => RotateFlipType.Rotate90FlipNone,
                8 => RotateFlipType.Rotate270FlipNone,
                _ => RotateFlipType.RotateNoneFlipNone,
            };
        }
        private static VectorOfVectorOfPoint FilterContours(VectorOfVectorOfPoint contours, double threshold = 100)
        {
            var cells = new List<Rectangle>();
            //filter out text contours by checking the size
            for (int i = 0; i < contours.Size; i++)
            {
                //get the area of the contour
                var area = CvInvoke.ContourArea(contours[i]);
                //filter out text contours using the area
                //if (area > 2000 && area < 200000)
                {
                    //check if the shape of the contour is a square or a rectangle
                    var rect = CvInvoke.BoundingRectangle(contours[i]);
                    var aspectRatio = (double)rect.Width / rect.Height;
                    if (aspectRatio > 0.5 && aspectRatio <= 5)
                    {
                        //add the cell to the list
                        cells.Add(rect);
                    }
                }
            }

            VectorOfVectorOfPoint? filteredContours = new VectorOfVectorOfPoint();
            for (int i = 0; i < contours.Size; i++)
            {


                if (CvInvoke.ContourArea(contours[i]) >= threshold)
                {
                    filteredContours.Push(contours[i]);
                }
            }

            return filteredContours;
        }
        private static List<System.Drawing.Rectangle> Contours2BBox(VectorOfVectorOfPoint contours)
        {
            List<System.Drawing.Rectangle> list = new List<System.Drawing.Rectangle>();
            for (int i = 0; i < contours.Size; i++)
            {
                if (contours[i].Length >= 3000)
                    list.Add(CvInvoke.BoundingRectangle(contours[i]));
            }

            return list;
        }
        private Mat GetHorizontalVertical(Mat ThresholdImage, int scale = 800)
        {
            MCvScalar mCvScalar = new MCvScalar(260);
            Point point = new Point(-1, -1);

            #region Horizontal lines            
            //Horizontal lines
            Mat? horizontal = ThresholdImage.Clone();
            int horizontalRow = horizontal.Rows / scale;
            int SizeXY = Convert.ToInt32(1);
            Mat? horizontalStructure = CvInvoke.GetStructuringElement(ElementShape.Rectangle, new Size(horizontalRow, SizeXY), point);
            CvInvoke.Erode(ThresholdImage, horizontal, horizontalStructure, new Point(-1, -1), -2, BorderType.Default, mCvScalar);
            CvInvoke.Dilate(horizontal, horizontal, horizontalStructure, new Point(-1, -1), 2, BorderType.Default, mCvScalar);
            #endregion

            #region Vertical lines
            Mat? vertical = ThresholdImage.Clone();
            int verticalCol = vertical.Cols / scale;
            Mat? verticalStructure = CvInvoke.GetStructuringElement(ElementShape.Rectangle, new Size(SizeXY, verticalCol), point);
            CvInvoke.Erode(ThresholdImage, vertical, verticalStructure, new Point(-1, -1), -2, BorderType.Default, mCvScalar);
            CvInvoke.Dilate(vertical, vertical, verticalStructure, new Point(-1, -1), Convert.ToInt32(1.5), BorderType.Default, mCvScalar);
            #endregion

            #region Both Horizontal & Vertical lines
            Mat? horizontalverticalMask = new Mat();
            CvInvoke.AddWeighted(horizontal, 10, vertical, 10, 10, horizontalverticalMask);

            Mat? bitxor = new Mat();
            CvInvoke.BitwiseXor(ThresholdImage, horizontalverticalMask, bitxor);
            CvInvoke.BitwiseAnd(bitxor, horizontalverticalMask, bitxor);
            CvInvoke.BitwiseOr(bitxor, horizontalverticalMask, bitxor);
            //CvInvoke.BitwiseNot(bitxor, horizontalverticalMask);

            return RemoveHorizontalVerticallines(ThresholdImage,horizontalverticalMask);
            #endregion


        }
        private Image<Gray, byte> SetBorder(Image<Gray, byte> mat)
        {
            Image<Gray, byte> image = mat;
            {
                int startCol = 1;
                int startRow = 1;
                int endCol = image.Cols;
                int endRow = image.Rows;
                Rectangle roi = new Rectangle(startCol, startRow, endCol - startCol, endRow - startRow);
                image.Draw(roi, new Gray(150), 30, LineType.EightConnected);
            }
            return image;
        }

        private Mat RemoveHorizontalVerticallines(Mat ThreshlgImage, Mat horizontalverticalMask)
        {
            Mat horizontalverticalMask1 = horizontalverticalMask;
            Mat mat = GrayscaleToBinaryThreshold(ThreshlgImage, 255, 255, true);
            CvInvoke.Threshold(mat, horizontalverticalMask, 2, 255, Emgu.CV.CvEnum.ThresholdType.BinaryInv | Emgu.CV.CvEnum.ThresholdType.Otsu);
            Mat edges = mat.Clone();
            int thickness = Convert.ToInt32(1);
            VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();
            CvInvoke.FindContours(horizontalverticalMask1, contours, null, RetrType.Tree, ChainApproxMethod.ChainApproxNone);
            
            List<System.Drawing.Rectangle> lisRectangle = Contours2BBox(FilterContours(contours));
            if (lisRectangle.Count > 0)
            {
                int imageWidth = edges.Cols;
                int imageHight = edges.Rows;
                List<System.Drawing.Rectangle> sortedBBoxes = lisRectangle.OrderBy(x => x.Y).ThenBy(y => y.X).ToList();
                int icoutnt = 0;
                int offset = 2;
                try
                {
                    for (int key = 0; key < sortedBBoxes.Count; key++)
                    {
                        icoutnt = key;

                        var rect = sortedBBoxes[key];

                        rect.X = (rect.X+ offset);
                        rect.Y = (rect.Y+ offset);
                        rect.Width = (rect.Width +offset);
                        rect.Height = (rect.Height + offset);
                        if (rect.Height < 0) continue;
                        if (rect.X > 0 && rect.Y > 0 && rect.Width > 150 && rect.Height > 15)
                        {
                            Image<Gray, byte> image1 = OrignalImage.ToImage<Gray, byte>();
                            image1.ROI = rect;

                            CvInvoke.Threshold(image1, image1, 2, 255, Emgu.CV.CvEnum.ThresholdType.BinaryInv | Emgu.CV.CvEnum.ThresholdType.Otsu);
                            image1 = image1.SmoothGaussian(1, 1, 5, 5);//1, 1, 5, 5
                            image1 = image1.SmoothBilateral(5, 5, 5);
                            image1 = image1.SmoothBlur(1, 1,true);

                            string dirpath = Path.Combine(@"D:\Channel\Code Sell\Working\Output\");                                                                                 //
                            string filepath = Path.Combine(dirpath, key.ToString() + ".png");
                            image1.Save(filepath);
                            image1 = null;
                        }
                    }
                }
                catch
                {

                }
            }
            return edges;
        }

        private Mat GetLines(Mat ThreshlgImage)
        {
            LineSegment2D[] lines = CvInvoke.HoughLinesP(ThreshlgImage, 1, Math.PI / 180, 80, 150, 0);
            var horizontals = lines.Where(l => Math.Abs(l.P1.Y - l.P2.Y) <= 5).ToList();
            var verticals = lines.Where(l => Math.Abs(l.P1.X - l.P2.X) <= 5).ToList();
            horizontals = MergeHorizontalLines(horizontals);
            verticals = MergeHorizontalLines(verticals);

            foreach (var line in horizontals)
            {
                CvInvoke.Line(ThreshlgImage, line.P1, line.P2, new MCvScalar(0, 255, 0), 2);
            }
            foreach (var line in verticals)
            {
                CvInvoke.Line(ThreshlgImage, line.P1, line.P2, new MCvScalar(0, 0, 255), 2);
            }
            return ThreshlgImage;
        }
        List<LineSegment2D> MergeHorizontalLines(List<LineSegment2D> lines, int tolerance = 5, int maxGap = 15)
        {
            var merged = new List<LineSegment2D>();
            var processed = new HashSet<int>();

            for (int i = 0; i < lines.Count; i++)
            {
                if (processed.Contains(i)) continue;

                var line1 = lines[i];
                int yAvg = (line1.P1.Y + line1.P2.Y) / 2;
                int minX = Math.Min(line1.P1.X, line1.P2.X);
                int maxX = Math.Max(line1.P1.X, line1.P2.X);

                for (int j = i + 1; j < lines.Count; j++)
                {
                    if (processed.Contains(j)) continue;

                    var line2 = lines[j];
                    int y2Avg = (line2.P1.Y + line2.P2.Y) / 2;

                    if (Math.Abs(yAvg - y2Avg) <= tolerance)
                    {
                        int line2MinX = Math.Min(line2.P1.X, line2.P2.X);
                        int line2MaxX = Math.Max(line2.P1.X, line2.P2.X);

                        if (line2MinX <= maxX + maxGap && line2MaxX >= minX - maxGap)
                        {
                            // Update bounds
                            minX = Math.Min(minX, line2MinX   );
                            maxX = Math.Max(maxX, line2MaxX);
                            processed.Add(j);
                        }
                    }
                }

                var newLine = new LineSegment2D(new Point(minX, yAvg), new Point(maxX, yAvg));
                merged.Add(newLine);
            }

            return merged;
        }


        #region OCR
        private object? ConvertImageToTextOCR(string mFileName)
        {
            Image<Gray, byte> emguImageOCRGray;
            Image<Rgb, byte> GetOCRImage = new Image<Rgb, byte>(mFileName);
            emguImageOCRGray = GetOCRImage.Convert<Gray, byte>();
            byte[] bytes = emguImageOCRGray.Bytes.ToArray();
            return PerformOCRAsync(mFileName, datapath, bytes);
        }
        
        private object PerformOCRAsync(string imagePath, string ocddata, byte[] bytes)
        {
            OutputOCRText loutputOCRText = new OutputOCRText();
            using (var stream = new FileStream(imagePath, FileMode.Open))
            {
                using (var engine = new Tesseract.TesseractEngine(ocddata, "eng", Tesseract.EngineMode.TesseractAndLstm))
                {
                    engine.SetVariable("tessedit_pageseg_mode", "1");
                    using (var pix = Tesseract.Pix.LoadFromFile(imagePath))
                    {
                        engine.DefaultPageSegMode = Tesseract.PageSegMode.Auto;
                        using (var page = engine.Process(pix))
                        {
                            var text = page.GetText();
                            loutputOCRText.GetOCRTextWithROI = page.GetTsvText(1);
                            loutputOCRText.GetOCRExtractedHTML = page.GetHOCRText(1);
                            loutputOCRText.GetOCRExtractedUNLVText = page.GetUNLVText();
                            stringBuilder.AppendLine("********************************************************Start****************************************************************************");
                            stringBuilder.Append(text);
                            stringBuilder.AppendLine("********************************************************End****************************************************************************");
                           
                            return loutputOCRText;
                        }
                    }
                }
            }
        }

        #endregion
    }

    public class OutputOCRText
    {



        public string? GetOCRExtractedHTML { get; set; } = string.Empty;


        public string GetJsonString { get; set; } = string.Empty;


        public string GetXMLString { get; set; } = string.Empty;


        public string? GetOCRExtractedUNLVText { get; set; }

        public string GetOCRTextWithROI { get; set; } = string.Empty;


        public string? GetOCRExtractedUTF8Text { get; set; }
    }
}
