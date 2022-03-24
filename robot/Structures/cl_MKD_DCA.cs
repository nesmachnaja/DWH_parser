using System;

namespace robot
{
    class cl_MKD_DCA
    {
        private int _ln;
        private DateTime _payment_date;
        private string _dca_name;
        private double _payment_amount;
        private double _dca_comission_amount;
        private DateTime _reestr_date;

        public DateTime Reestr_date
        {
            get => _reestr_date;
            set => _reestr_date = value;
        }
        public int LN
        {
            get => _ln;
            set => _ln = value;
        }
        public DateTime Payment_date
        {
            get => _payment_date;
            set => _payment_date = value;
        }
        public string DCA_name
        {
            get => _dca_name;
            set => _dca_name = value;
        }
        public double Payment_amount
        {
            get => _payment_amount;
            set => _payment_amount = value;
        }
        public double DCA_comission_amount
        {
            get => _dca_comission_amount;
            set => _dca_comission_amount = value;
        }

    }
}
