using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Administracion_OMAJA;

namespace Administracion_OMAJA
{
    public class RegistroCarga
    {
        public DateTime FechaHora { get; set; }
        public int DocumentosCargados { get; set; }
        public int Nuevos { get; set; }
        public int Actualizados { get; set; }
    }
}
