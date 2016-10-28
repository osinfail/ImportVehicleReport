namespace ImportVehicleReport.Report
{
    abstract class Benefit
    {
        public int FtpSuccess { set; get; }
        public int FtpFailure { set; get; }

        public int ZipFiles { set; get; }
        public int ImportVehicleRecords { set; get; }
        public string NotWellFormedXmlCount { set; get; }
        public string NotFoundPos { set; get; }

        public int StockCount { set; get; }
        public int NewStockCount { set; get; }
        public int DeletedStockCount { set; get; }

        public Photo PhotoStatus { set; get; }

        protected Benefit()
        {
            FtpSuccess = 0;
            FtpFailure = 0;
            ZipFiles = 0;
            NotWellFormedXmlCount = string.Empty;
            NotFoundPos = string.Empty;
            StockCount = 0;
            NewStockCount = 0;
            DeletedStockCount = 0;
            PhotoStatus = new Photo();
        }
    }
}
