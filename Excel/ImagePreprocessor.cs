using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Text;

namespace OfficeFireSync.Excel
{
    public class ImagePreprocessor
    {
        public Image ResizeImage(string path, int size)
        {
            using var image = Image.FromFile(path);
            return ResizeImage(image, size);
        }

        public Image ResizeImage(Image image, int size) // Synchronous
        {
            if (image.Size.Width <= size && image.Size.Height <= size)
                return new Bitmap(image);

            var ratio = (float)Math.Min(image.Size.Width, image.Size.Height) / (float)Math.Max(image.Size.Width, image.Size.Height);
            var width = image.Size.Width >= image.Size.Height ? size : size * ratio;
            var height = image.Size.Height >= image.Size.Width ? size : size * ratio;

            using var bitmap = new Bitmap(image);
            Bitmap thumbBitmap = new Bitmap((int)width, (int)height);

            using Graphics g = Graphics.FromImage(thumbBitmap);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.DrawImage(bitmap, 0, 0, width, height);
            return thumbBitmap;
        }

        public Image ResizeImage(byte[] binary, int size)
        {
            using var memStream = new MemoryStream(binary);
            Image image = Image.FromStream(memStream);
            return ResizeImage(image, size);
        }

        public string ConvertToBase64(string path)
        {
            using var image = Image.FromFile(path);
            return ConvertToBase64(image);
        }

        public string ConvertToBase64(Image image)
        {
            using var memStream = new MemoryStream();
            image.Save(memStream, ImageFormat.Jpeg);
            memStream.Position = 0;
            return ConvertToBase64(memStream.ToArray(), GetFormat(image));
        }

        public string ConvertToBase64(byte[] binary, string ext = "jpg")
        {
            var base64Data = Convert.ToBase64String(binary);
            return $"data:image/{ext};base64,{base64Data}";
        }

        public string GetFormat(Image image)
        {
            if (ImageFormat.Jpeg.Equals(image.RawFormat))
            {
                return "jpg";
            }
            else if (ImageFormat.Png.Equals(image.RawFormat))
            {
                return "png";
            }
            else if (ImageFormat.Gif.Equals(image.RawFormat))
            {
                return "gif";
            } else
            {
                return "bmp";
            }

            //throw new NotImplementedException("Image format not supported!");
        }
    }
}
