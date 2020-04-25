using System;
using System.Drawing;
using System.Drawing.Imaging;
using Svg;
using System.Xml;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;

public partial class UserDefinedFunctions
{

    [SqlFunction]
    public static String SVGToPNGByteArray(String inputstring)
    {

        var svgContent = inputstring;
        var byteArray = Convert.FromBase64String(svgContent);

        using (var stream = new MemoryStream(byteArray))
        {

            var svgDocument = SvgDocument.Open<SvgDocument>(stream);
            var stream1 = new MemoryStream();

            int W = (int)svgDocument.Width.Value;
            int H = (int)svgDocument.Height.Value;

            Bitmap bmp = svgDocument.Draw(W, H);

            bmp.Save(stream1, ImageFormat.Png);// save Bitmap as PNG-File
            byte[] bytes = stream1.ToArray();
            var output = Convert.ToBase64String(bytes);
            return output;

        }
    }
}
