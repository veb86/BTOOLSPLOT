using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BTOOLS_PLOT
{
    public class settingXML
    {
		public struct setSheet
		{
			public int height;
			public int weight;
			public string printPortrait;
			public string printLandscape;
			public string nameSheet;
			public int weightPlot;
			public int heightPlot;

		}

		public string waysavetopdf;
		public string waysavetoext;
		public string colorWB;
		public string colorRGB;
		public string specTextColor;
        public string prefixSheet;
        public string suffixSheet;
        public string connectorSheet;
        public bool swapNameAndNum;

        public List<setSheet> listFS = new List<setSheet>();

		//public List<Double> arrScale = new List<Double>();
		//public List<Double> listLengthLine = new List<Double>();
		//public List<Double> listOrigLine = new List<Double>();

	}
}
