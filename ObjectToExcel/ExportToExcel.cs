using System;

namespace ObjectToExcel
{
    public class ExportToExcel : Attribute
    {
        private int Order;
        private bool Wrap;

        public ExportToExcel(int order, bool wrap = true)
        {
            this.Order = order;
            this.Wrap = wrap;
        }

        public virtual int order
        {
            get { return this.order; }
        }
        public virtual bool wrap
        {
            get { return this.Wrap; }
        }
    }
}