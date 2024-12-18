using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class ImageDefine
    {
        public ImageDefine(string imageUrl, List<ImageInfo> imageInfos)
        {
            ImageUrl = imageUrl;
            ImageInfos = imageInfos;
        }

        public string ImageUrl { get; set; }

        public virtual List<ImageInfo> ImageInfos { get; set; }
    }

    public class ImageInfo
    {
        public ImageInfo(string imageName, double px, double py)
        {
            ImageName = imageName;
            Px = px;
            Py = py;
        }

        public string ImageName { get; set; }

        public double Px { get; set; }

        public double Py { get; set; }
    }
}
