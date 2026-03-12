using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace ProyectoGRE.DTO
{
    [Serializable]
    public class DtoB : ClassResultPV, INotifyPropertyChanged
    {

        private string msjError;
        private string sqlQuery;

        [Browsable(false)]
        public string MsjError
        {
            get { return msjError; }
            set { msjError = value; }
        }

        [Browsable(false)]
        public DtoB Error(string msj)
        {
            msjError = msj;
            return this;
        }

        [Browsable(false)]
        public string SqlQuery
        {
            get { return sqlQuery; }
            set { sqlQuery = value; }
        }

 

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

    }
}
