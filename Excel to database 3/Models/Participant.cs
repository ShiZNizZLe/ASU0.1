using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_to_database_3.Models
{
    class Participant
    {
        private int BSN { get; set; }
        private int LKN { get; set; }
        private string Naam { get; set; }
        private DateTime GeboorteDatum { get; set; }
        private DateTime EersteZiektedag { get; set; }
        private DateTime HerstelDatum { get; set; }
        private DateTime ZiekteMeldingOntvangen { get; set; }
        private string AOKlasse { get; set; }
        private int AOPercentage { get; set; }
        private DateTime ActieDatum { get; set; }
        private string ActieInhoud { get; set; }
    }
}
