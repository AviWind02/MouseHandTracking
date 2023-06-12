using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Point = OpenCvSharp.Point;
using Rect = OpenCvSharp.Rect;
using Size = OpenCvSharp.Size;
using Window = System.Windows.Window;

namespace MouseHandTracking
{
    public partial class MainWindow : Window
    {
        private VideoCapture _videoCapture;
        private DispatcherTimer _timer;

        public MainWindow()
        {
            InitializeComponent();

            // Initialize the VideoCapture object with the default webcam
            _videoCapture = new VideoCapture(0);

            // Set the maximum buffer size to limit the frame backlog
            _videoCapture.Set(VideoCaptureProperties.OPENNI_MaxBufferSize, 1);

            // Create a timer to update the image
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromMilliseconds(8); // Update at approximately 30 frames per second
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        private BitmapSource MatToBitmapSource(Mat mat)
        {
            try
            {
                // Convert the Mat to a byte array
                byte[] imageData = new byte[mat.Rows * mat.Cols * mat.ElemSize()];
                Marshal.Copy(mat.Data, imageData, 0, imageData.Length);

                // Create a BitmapSource
                var bitmapSource = BitmapSource.Create(mat.Width, mat.Height, 96, 96,
                    mat.Channels() == 3 ? PixelFormats.Bgr24 : PixelFormats.Gray8,
                    null, imageData, (int)mat.Step());

                return bitmapSource;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to convert Mat to BitmapSource: " + ex.Message);
                return null;
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            // Add a small delay to retrieve frames at a more consistent rate
            System.Threading.Thread.Sleep(30);

            // Read a frame from the webcam
            Mat frame = new Mat();
            _videoCapture.Read(frame);

            // Check if the frame is valid
            if (frame.Empty())
                return;

            // Resize the frame to a smaller size for smoother performance
            Mat resizedFrame = new Mat();
            Cv2.Resize(frame, resizedFrame, new Size(640, 480)); // Adjust the size as needed

            // Convert the OpenCV Mat to a BitmapSource
            BitmapSource bitmapSource = MatToBitmapSource(resizedFrame);

            // Update the image control with the new frame
            image.Source = bitmapSource;

            // Hand detection and console logging
            DetectAndLogHand(resizedFrame);
        }

        private void DetectAndLogHand(Mat frame)
        {
            // Convert the frame to grayscale for better hand detection
            Mat grayFrame = new Mat();
            Cv2.CvtColor(frame, grayFrame, ColorConversionCodes.BGR2GRAY);

            // Apply Gaussian blur to reduce noise
            Cv2.GaussianBlur(grayFrame, grayFrame, new Size(7, 7), 0);

            // Apply thresholding to segment the hand from the background
            Cv2.Threshold(grayFrame, grayFrame, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);

            // Find contours in the binary image
            Point[][] contours;
            HierarchyIndex[] hierarchy;
            Cv2.FindContours(grayFrame, out contours, out hierarchy, RetrievalModes.External, ContourApproximationModes.ApproxSimple);

            // Find the largest contour (assumed to be the hand)
            double maxArea = 0;
            int maxAreaIdx = -1;
            for (int i = 0; i < contours.Length; i++)
            {
                double area = Cv2.ContourArea(contours[i]);
                if (area > maxArea)
                {
                    maxArea = area;
                    maxAreaIdx = i;
                }
            }

            // Check if a hand contour is found
            if (maxAreaIdx >= 0)
            {
                // Get the bounding rectangle of the hand contour
                Rect handRect = Cv2.BoundingRect(contours[maxAreaIdx]);

                // Log hand information
                Console.WriteLine("Hand Detected");
                Console.WriteLine("  - Hand Position (X, Y): " + handRect.X + ", " + handRect.Y);
                Console.WriteLine("  - Hand Size (Width, Height): " + handRect.Width + ", " + handRect.Height);

                // Draw the hand contour and bounding rectangle on the frame
                Cv2.DrawContours(frame, contours, maxAreaIdx, Scalar.Red, 2);
                Cv2.Rectangle(frame, handRect, Scalar.Blue, 2);
            }
        }

    }
}
