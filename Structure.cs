namespace PYVS_CCMDataCollector
{
    partial class Program
    {
        //CCM operation data
        public struct CCMData
        {
            public string heat;            //heat number   
            public string porder;          //pre-order number
            public short seq;              //sequence number (default = 1)

            public double moldwaterInletStd1A;   //Mold Water Inlet average temperature A line strand 1
            public double moldwaterInletStd1B;   //Mold Water Inlet average temperature B line strand 1
            public double moldwaterInletStd2A;   //Mold Water Inlet average temperature A line strand 2
            public double moldwaterInletStd2B;   //Mold Water Inlet average temperature B line strand 2
            public double moldwaterInletStd3A;   //Mold Water Inlet average temperature A line strand 3
            public double moldwaterInletStd3B;   //Mold Water Inlet average temperature B line strand 3
            public double moldwaterInletStd4A;   //Mold Water Inlet average temperature A line strand 4
            public double moldwaterInletStd4B;   //Mold Water Inlet average temperature B line strand 4
            public double moldwaterInletStd5A;   //Mold Water Inlet average temperature A line strand 5
            public double moldwaterInletStd5B;   //Mold Water Inlet average temperature B line strand 5
            public double moldwaterInletStd6A;   //Mold Water Inlet average temperature A line strand 6
            public double moldwaterInletStd6B;   //Mold Water Inlet average temperature B line strand 6


            public void InitData()
            {
                this.heat = null;
                this.porder = null;
                this.seq = 0;

                this.moldwaterInletStd1A = 0;
                this.moldwaterInletStd1B = 0;
                this.moldwaterInletStd2A = 0;
                this.moldwaterInletStd2B = 0;
                this.moldwaterInletStd3A = 0;
                this.moldwaterInletStd3B = 0;
                this.moldwaterInletStd4A = 0;
                this.moldwaterInletStd4B = 0;
                this.moldwaterInletStd5A = 0;
                this.moldwaterInletStd5B = 0;
                this.moldwaterInletStd6A = 0;
                this.moldwaterInletStd6B = 0;
            }
        }
    }
}
