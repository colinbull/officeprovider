namespace EonUk.Retail.Ice.NonSupplyCustomerServicingModule.Helpers.Office.Excel
{
    public class ExcelPagedArray
    {
        private int? _pageSize;

        public object[][] Data { get; set; }

        public int PageSize
        {
            get
            {
                if (this._pageSize.HasValue)
                {
                    return this._pageSize.Value;
                }
                
                return this.Data.GetLength(0);
            }
            set { this._pageSize = value; }
        }

        public uint Length
        {
            get { return (uint)this.Data.GetLength(0); }
        }

        public object[] this[int i]
        {
            get { return this.Data[i]; }
            set { this.Data[i] = value; }
        }
    }
}