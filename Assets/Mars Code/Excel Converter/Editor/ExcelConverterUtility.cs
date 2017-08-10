
namespace MarsCode.ExcelConverter
{
    using UnityEditor;
    using UnityEngine;

    public static class ExcelConverterUtility
    {

        public static Color ExcelStandardColor { get { return new Color32(2, 113, 45, 255); } }

        public static Color StandardRed { get { return new Color32(255, 102, 102, 255); } }

        public static Color StandardBlue { get { return new Color32(153, 204, 255, 255); } }

        public static Color StandardYellow { get { return new Color32(255, 212, 128, 255); } }


        public static Texture2D GenerateColorBlockTexture(Color col, int width = 1, int height = 1)
        {
            var pix = new Color[width * height];
            for(var i = 0; i < pix.Length; i++)
            {
                pix[i] = col;
            }

            var output = new Texture2D(width, height);
            output.SetPixels(pix);
            output.Apply();

            return output;
        }

    }
}
