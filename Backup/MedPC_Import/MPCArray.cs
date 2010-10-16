namespace MedPC_Import
{
    struct MPCArray
    {
        public string name;
        public string outputName;
        public string summary;
        public string outputStyle;
        public System.Collections.ArrayList columns;
    }

    struct MPCArrayColumn
    {
        public string name;
        public string outputName;
        public string summary;
        public int outputColNum;
        public bool includeInOutput;
    }
}