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

            public double moldwaterOutletPressureStd1A;  //Mold water outlet average pressure A line strand 1
            public double moldwaterOutletPressureStd1B;  //Mold water outlet average pressure B line strand 1
            public double moldwaterOutletPressureStd2A;  //Mold water outlet average pressure A line strand 2
            public double moldwaterOutletPressureStd2B;  //Mold water outlet average pressure B line strand 2
            public double moldwaterOutletPressureStd3A;  //Mold water outlet average pressure A line strand 3
            public double moldwaterOutletPressureStd3B;  //Mold water outlet average pressure B line strand 3
            public double moldwaterOutletPressureStd4A;  //Mold water outlet average pressure A line strand 4
            public double moldwaterOutletPressureStd4B;  //Mold water outlet average pressure B line strand 4
            public double moldwaterOutletPressureStd5A;  //Mold water outlet average pressure A line strand 5
            public double moldwaterOutletPressureStd5B;  //Mold water outlet average pressure B line strand 5
            public double moldwaterOutletPressureStd6A;  //Mold water outlet average pressure A line strand 6
            public double moldwaterOutletPressureStd6B;  //Mold water outlet average pressure B line strand 6

            public double castSpeedMinStrand1;  //Cast Speed minimum of strand 1
            public double castSpeedAvgStrand1;  //Cast Speed average of strnad 1
            public double castSpeedMaxStrand1;  //Cast Speed maximum of strand 1
            public double castSpeedMinStrand2;  //Cast Speed minimum of strand 2
            public double castSpeedAvgStrand2;  //Cast Speed average of strnad 2
            public double castSpeedMaxStrand2;  //Cast Speed maximum of strand 2
            public double castSpeedMinStrand3;  //Cast Speed minimum of strand 3
            public double castSpeedAvgStrand3;  //Cast Speed average of strnad 3
            public double castSpeedMaxStrand3;  //Cast Speed maximum of strand 3
            public double castSpeedMinStrand4;  //Cast Speed minimum of strand 4
            public double castSpeedAvgStrand4;  //Cast Speed average of strnad 4
            public double castSpeedMaxStrand4;  //Cast Speed maximum of strand 4
            public double castSpeedMinStrand5;  //Cast Speed minimum of strand 5
            public double castSpeedAvgStrand5;  //Cast Speed average of strnad 5
            public double castSpeedMaxStrand5;  //Cast Speed maximum of strand 5
            public double castSpeedMinStrand6;  //Cast Speed minimum of strand 6
            public double castSpeedAvgStrand6;  //Cast Speed average of strnad 6
            public double castSpeedMaxStrand6;  //Cast Speed maximum of strand 6

            public double mouldLevelMinStrand1; //Mould level minimum of strand 1
            public double mouldLevelAvgStrand1; //Mould level average of strand 1
            public double mouldLevelMaxStrand1; //Mould level average of strand 1
            public double mouldLevelMinStrand2; //Mould level minimum of strand 2
            public double mouldLevelAvgStrand2; //Mould level average of strand 2
            public double mouldLevelMaxStrand2; //Mould level average of strand 2
            public double mouldLevelMinStrand3; //Mould level minimum of strand 3
            public double mouldLevelAvgStrand3; //Mould level average of strand 3
            public double mouldLevelMaxStrand3; //Mould level average of strand 3
            public double mouldLevelMinStrand4; //Mould level minimum of strand 4
            public double mouldLevelAvgStrand4; //Mould level average of strand 4
            public double mouldLevelMaxStrand4; //Mould level average of strand 4
            public double mouldLevelMinStrand5; //Mould level minimum of strand 5
            public double mouldLevelAvgStrand5; //Mould level average of strand 5
            public double mouldLevelMaxStrand5; //Mould level average of strand 5
            public double mouldLevelMinStrand6; //Mould level minimum of strand 6
            public double mouldLevelAvgStrand6; //Mould level average of strand 6
            public double mouldLevelMaxStrand6; //Mould level average of strand 6


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

                this.moldwaterOutletPressureStd1A = 0;
                this.moldwaterOutletPressureStd1B = 0;
                this.moldwaterOutletPressureStd2A = 0;
                this.moldwaterOutletPressureStd2B = 0;
                this.moldwaterOutletPressureStd3A = 0;
                this.moldwaterOutletPressureStd3B = 0;
                this.moldwaterOutletPressureStd4A = 0;
                this.moldwaterOutletPressureStd4B = 0;
                this.moldwaterOutletPressureStd5A = 0;
                this.moldwaterOutletPressureStd5B = 0;
                this.moldwaterOutletPressureStd6A = 0;
                this.moldwaterOutletPressureStd6B = 0;

                this.castSpeedMinStrand1 = 0;
                this.castSpeedAvgStrand1 = 0;  //Cast Speed average of strnad 1
                this.castSpeedMaxStrand1 = 0;  //Cast Speed maximum of strand 1
                this.castSpeedMinStrand2 = 0;  //Cast Speed minimum of strand 2
                this.castSpeedAvgStrand2 = 0;  //Cast Speed average of strnad 2
                this.castSpeedMaxStrand2 = 0;  //Cast Speed maximum of strand 2
                this.castSpeedMinStrand3 = 0;  //Cast Speed minimum of strand 3
                this.castSpeedAvgStrand3 = 0;  //Cast Speed average of strnad 3
                this.castSpeedMaxStrand3 = 0;  //Cast Speed maximum of strand 3
                this.castSpeedMinStrand4 = 0;  //Cast Speed minimum of strand 4
                this.castSpeedAvgStrand4 = 0;  //Cast Speed average of strnad 4
                this.castSpeedMaxStrand4 = 0;  //Cast Speed maximum of strand 4
                this.castSpeedMinStrand5 = 0;  //Cast Speed minimum of strand 5
                this.castSpeedAvgStrand5 = 0;  //Cast Speed average of strnad 5
                this.castSpeedMaxStrand5 = 0;  //Cast Speed maximum of strand 5
                this.castSpeedMinStrand6 = 0;  //Cast Speed minimum of strand 6
                this.castSpeedAvgStrand6 = 0;  //Cast Speed average of strnad 6
                this.castSpeedMaxStrand6 = 0;  //Cast Speed maximum of strand 6

                this.mouldLevelMinStrand1 = 0; //Mould level minimum of strand 1
                this.mouldLevelAvgStrand1 = 0; //Mould level average of strand 1
                this.mouldLevelMaxStrand1 = 0; //Mould level average of strand 1
                this.mouldLevelMinStrand2 = 0; //Mould level minimum of strand 2
                this.mouldLevelAvgStrand2 = 0; //Mould level average of strand 2
                this.mouldLevelMaxStrand2 = 0; //Mould level average of strand 2
                this.mouldLevelMinStrand3 = 0; //Mould level minimum of strand 3
                this.mouldLevelAvgStrand3 = 0; //Mould level average of strand 3
                this.mouldLevelMaxStrand3 = 0; //Mould level average of strand 3
                this.mouldLevelMinStrand4 = 0; //Mould level minimum of strand 4
                this.mouldLevelAvgStrand4 = 0; //Mould level average of strand 4
                this.mouldLevelMaxStrand4 = 0; //Mould level average of strand 4
                this.mouldLevelMinStrand5 = 0; //Mould level minimum of strand 5
                this.mouldLevelAvgStrand5 = 0; //Mould level average of strand 5
                this.mouldLevelMaxStrand5 = 0; //Mould level average of strand 5
                this.mouldLevelMinStrand6 = 0; //Mould level minimum of strand 6
                this.mouldLevelAvgStrand6 = 0; //Mould level average of strand 6
                this.mouldLevelMaxStrand6 = 0; //Mould level average of strand 6
        }
        }

        public struct TundishData
        {
            public string heat;             //heat number (Tundish heat number)
            public double tundishWeight;    //tundish weight
            public double tundishHeight;    //tundish vertical position (Auto shifting actual position)

            public void InitData()
            {
                heat = null;
                tundishWeight = 0;
                tundishHeight = 0;
            }
        }

        public struct CutLengthData
        {
            public string heat;
            public short sequenceNo;
            public short strandNo;
            public short scaleNo;

            public double lengthTarget;
            public double lengthLastcut;
            public short lengthCompensation;
            public double weight;
            public double lengthMeasured;

            public void InitData()
            {
                heat=null;
                sequenceNo = 0;
                strandNo = 0;
                scaleNo = 0;

                lengthTarget = 0;
                lengthLastcut = 0;
                lengthCompensation = 0;
                weight = 0;
                lengthMeasured = 0;
            }
        }
    }
}
