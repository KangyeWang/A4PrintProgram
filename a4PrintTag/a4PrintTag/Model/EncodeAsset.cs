using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test2.Model
{
    public partial class EncodeAsset
    {
        public string AssetName
        {
            get;
            set;
        }

        public string AssetDate
        {
            set;
            get;
        }

        public string AssetYear
        {
            set;
            get;
        }

        public Image EncodedErp
        {
            set;
            get;
        }
    }
}
