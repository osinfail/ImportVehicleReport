namespace ImportVehicleReport.Report
{
    class Report
    {
        public PlanetVo PlanetVoStatus { set; get; }
        public Tec3H Tec3HStatus { set; get; }

        public bool HavasStatus { set; get; }

        public string PdvNameChange { set; get; }
        public bool LuceneStatus { set; get; }
        public bool ResetApplicationPoolStatus { set; get; }
        public bool ImportVehicleStatus { set; get; }
        public bool XmlPdvFile { set; get; }

        public Report()
        {
            PlanetVoStatus = new PlanetVo();
            Tec3HStatus = new Tec3H();

            HavasStatus = false;

            PdvNameChange = string.Empty;
            LuceneStatus = false;
            ImportVehicleStatus = false;
            XmlPdvFile = false;
        }
    }
}
