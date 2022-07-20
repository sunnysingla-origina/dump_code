using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace CaptureMostRecentLiDocumentForIBM
{
    public class ConfigClass
    {
        public ConfigClass()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        public static string con
        {
            get
            {

                if (ConfigurationManager.ConnectionStrings["connectionstring"].ConnectionString.ToString() != null)
                    return ConfigurationManager.ConnectionStrings["connectionstring"].ConnectionString.ToString();
                else
                    return "";
            }
        }

        
      

    }

}
