using DocumentFormat.OpenXml.Packaging;
using OfficeFireSync.Excel;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeFireSync.Word
{
    public class HTMLWordETL : WordETL
    {
        private readonly ImagePreprocessor imagePreprocessor;
        public HTMLWordETL(ImagePreprocessor imagePreprocessor) : base() 
        {
            this.imagePreprocessor = imagePreprocessor;
        }

        public void Test(string filePath)
        {
            byte[] byteArray = File.ReadAllBytes(filePath);
            using MemoryStream memoryStream = new MemoryStream();
            memoryStream.Write(byteArray, 0, byteArray.Length);
            using WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true);

            int imageCounter = 0;
            HtmlConverterSettings settings = new HtmlConverterSettings()
            {
                PageTitle = "My Page Title",
                ImageHandler = imageInfo => {
                    // This property dictates how images are handled
                    ++imageCounter;
                    string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                    ImageFormat imageFormat = null;
                    if (extension == "png") imageFormat = ImageFormat.Png;
                    else if (extension == "gif") imageFormat = ImageFormat.Gif;
                    else if (extension == "bmp") imageFormat = ImageFormat.Bmp;
                    else if (extension == "jpeg") imageFormat = ImageFormat.Jpeg;
                    else if (extension == "tiff")
                    {
                        extension = "gif";
                        imageFormat = ImageFormat.Gif;
                    }
                    else if (extension == "x-wmf")
                    {
                        extension = "wmf";
                        imageFormat = ImageFormat.Wmf;
                    }

                    if (imageFormat == null) return null;

                    string base64 = null;

                    try
                    {
                        // Read the image and converts it to Base64String
                        using (MemoryStream ms = new MemoryStream())
                        {
                            imageInfo.Bitmap.Save(ms, imageFormat);
                            var ba = ms.ToArray();
                            var image = imagePreprocessor.ResizeImage(ba, 150);
                            using (MemoryStream ms2 = new MemoryStream())
                            {
                                image.Save(ms2, imageFormat);
                                ba = ms2.ToArray();
                                //base64 = imagePreprocessor.ConvertToBase64(image);
                                base64 = Convert.ToBase64String(ba);
                            }
                        }
                    }
                    catch (System.Runtime.InteropServices.ExternalException)
                    { return null; }

                    ImageFormat format = imageInfo.Bitmap.RawFormat;
                    ImageCodecInfo codec = ImageCodecInfo.GetImageDecoders()
                                              .First(c => c.FormatID == format.Guid);
                    string mimeType = codec.MimeType;

                    string imageSource =
                           string.Format("data:{0};base64,{1}", mimeType, base64);

                    XElement img = new XElement(Xhtml.img,
                          new XAttribute(NoNamespace.src, imageSource),
                          imageInfo.ImgStyleAttribute,
                          imageInfo.AltText != null ?
                               new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                    return img;
                }
            };
            XElement html = HtmlConverter.ConvertToHtml(doc, settings);
            var htmlString = html.ToString(SaveOptions.DisableFormatting);
            File.WriteAllText(@"C:\Users\micha\Desktop\test.html", htmlString, Encoding.UTF8);
            //File.WriteAllText(@"C:\Users\micha\Desktop\test.html", html.ToStringNewLineOnAttributes());
        }
    }
}
