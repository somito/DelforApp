using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.IO;

namespace DelforApp
{
    public class Delfor
    {
        public Delfor(string path)
        {
            DelforMessageXml = XElement.Load(path);
            FileName = Path.GetFileName(path);
        }

        public XElement DelforMessageXml { get; set; }
        public string FileName { get; set; }
        public string MessageFrom { get; set; }
        public string MessageTo { get; set; }
        public string CustomerPlant { get; set; }
        public string AliasNumber { get; set; }
        public string Contract { get; set; }
        public string LoadingPlace { get; set; }
        public string ShipToPlace { get; set; }
        public string Warehouse { get; set; }
        public string BuyerID { get; set; }
        public string SellerID { get; set; }
        public string PlantCode { get; set; }
        public string Loc1 { get; set; }
        public string Loc2 { get; set; }
        public string VWType10 { get; set; }
        public string TimeOfMessage { get; set; }

        XNamespace ns = "http://www.intentia.com/Schemas/mec";
        XNamespace nsenv = "http://www.intentia.com/MBM_Envelope_1";
        XNamespace nsxsi = "http://www.w3.org/2001/XMLSchema-instance";
        XNamespace nsxsd = "http://www.w3.org/2001/XMLSchema";


        public string GetMessageFrom(XElement xml)
        {
            var from = xml.Descendants()
              .Where(x => x.Name.LocalName == "from")
              .Select(x => x.Element(nsenv + "address").Value);

            return string.Join("", from);
        }

        public string GetMessageTo(XElement xml)
        {
            var to = xml.Descendants()
              .Where(x => x.Name.LocalName == "to")
              .Select(x => x.Element(nsenv + "address").Value);

            return string.Join("", to);
        }

        public string GetAliasNumber(XElement xml)
        {
            var alias = xml.Descendants()
              .Where(x => x.Name.LocalName == "LIN")
              .Descendants()
              .Where(x => x.Name.LocalName == "cmp01")
              .Select(x => x.Element(ns + "e01_7140").Value);

            return string.Join(";", alias);
        }

        public string GetBuyerID(XElement xml)
        {
            var type = xml.Descendants()
                       .Where(x => x.Name.LocalName == "NAD")
                       .Select(x => x.Element(ns + "e01_3035").DescendantsAndSelf());

            for (int i = 0; i < type.Count(); i++)
            {
                if (string.Join("", type.ElementAt(i).ElementAt(0).Value) == "BY")
                {
                    var ID = type.ElementAt(i).ElementAt(0).ElementsAfterSelf()
                             .Where(x => x.Name.LocalName == "cmp01")
                             .Select(x => x.Element(ns + "e01_3039").Value);

                    return string.Join("", ID);
                }
            }

            return null;

        }

        public string GetSellerID(XElement xml)
        {
            var type = xml.Descendants()
                       .Where(x => x.Name.LocalName == "NAD")
                       .Select(x => x.Element(ns + "e01_3035").DescendantsAndSelf());

            for (int i = 0; i < type.Count(); i++)
            {
                if (string.Join("", type.ElementAt(i).ElementAt(0).Value) == "SE")
                {
                    var ID = type.ElementAt(i).ElementAt(0).ElementsAfterSelf()
                             .Where(x => x.Name.LocalName == "cmp01")
                             .Select(x => x.Element(ns + "e01_3039").Value);

                    return string.Join("", ID);
                }

                else if (string.Join("", type.ElementAt(i).ElementAt(0).Value) == "SU")
                {
                    var ID = type.ElementAt(i).ElementAt(0).ElementsAfterSelf()
                             .Where(x => x.Name.LocalName == "cmp01")
                             .Select(x => x.Element(ns + "e01_3039").Value);

                    return string.Join("", ID);
                }
            }

            return null;

        }

        public string GetLoadingPlace(XElement xml)
        {
            var type = xml.Descendants()
                       .Where(x => x.Name.LocalName == "NAD")
                       .Select(x => x.Element(ns + "e01_3035").DescendantsAndSelf());

            for (int i = 0; i < type.Count(); i++)
            {
                if (string.Join("", type.ElementAt(i).ElementAt(0).Value) == "SF")
                {
                    var ID = type.ElementAt(i).ElementAt(0).ElementsAfterSelf()
                             .Where(x => x.Name.LocalName == "cmp03")
                             .Select(x => x.Element(ns + "e01_3036").Value);

                    return string.Join("", ID);
                }
            }

            return null;

        }

        public string GetShipToPlace(XElement xml)
        {
            var type = xml.Descendants()
                       .Where(x => x.Name.LocalName == "NAD")
                       .Select(x => x.Element(ns + "e01_3035").DescendantsAndSelf());

            for (int i = 0; i < type.Count(); i++)
            {
                if (string.Join("", type.ElementAt(i).ElementAt(0).Value) == "ST")
                {
                    var ID = type.ElementAt(i).ElementAt(0).ElementsAfterSelf()
                             .Where(x => x.Name.LocalName == "cmp03")
                             .Select(x => x.Element(ns + "e01_3036").Value);

                    return string.Join("", ID);
                }
            }

            return null;

        }

        public string GetPlantCode(XElement xml)
        {
            var type = xml.Descendants()
                       .Where(x => x.Name.LocalName == "NAD")
                       .Select(x => x.Element(ns + "e01_3035").DescendantsAndSelf());

            for (int i = 0; i < type.Count(); i++)
            {
                if (string.Join("", type.ElementAt(i).ElementAt(0).Value) == "ST")
                {
                    var ID = type.ElementAt(i).ElementAt(0).ElementsAfterSelf()
                             .Where(x => x.Name.LocalName == "cmp01")
                             .Select(x => x.Element(ns + "e01_3039").Value);

                    return string.Join("", ID);
                }
            }

            return null;
        }

        public string GetLoc1(XElement xml)
        {
            var type = xml.Descendants()
                       .Where(x => x.Name.LocalName == "LOC")
                       .Select(x => x.Element(ns + "e01_3227").DescendantsAndSelf());

            for (int i = 0; i < type.Count(); i++)
            {
                if (string.Join("", type.ElementAt(i).ElementAt(0).Value) == "11")
                {
                    var ID = type.ElementAt(i).ElementAt(0).ElementsAfterSelf()
                             .Where(x => x.Name.LocalName == "cmp01")
                             .Select(x => x.Element(ns + "e01_3225").Value);

                    return string.Join("", ID);
                }
            }

            return null;
        }

        public string GetLoc2(XElement xml)
        {
            var type = xml.Descendants()
                       .Where(x => x.Name.LocalName == "LOC")
                       .Select(x => x.Element(ns + "e01_3227").DescendantsAndSelf());

            for (int i = 0; i < type.Count(); i++)
            {
                if (string.Join("", type.ElementAt(i).ElementAt(0).Value) == "7")
                {
                    var ID = type.ElementAt(i).ElementAt(0).ElementsAfterSelf()
                             .Where(x => x.Name.LocalName == "cmp01")
                             .Select(x => x.Element(ns + "e01_3225").Value);

                    return string.Join("", ID);
                }
            }

            return null;
        }

        public string GetTimeOfMessage(XElement xml)
        {
            var ymd = xml.Descendants()
              .Where(x => x.Name.LocalName == "UNB")
              .Descendants()
              .Where(x => x.Name.LocalName == "cmp04")
              .Select(x => x.Element(ns + "e01_0017").Value);

            var ymdstring = string.Join("", ymd);

            var day = ymdstring.Substring(4,2);
            var month = ymdstring.Substring(2, 2);
            var year = ymdstring.Substring(0, 2);

            var hm = xml.Descendants()
              .Where(x => x.Name.LocalName == "UNB")
              .Descendants()
              .Where(x => x.Name.LocalName == "cmp04")
              .Select(x => x.Element(ns + "e02_0019").Value);

            var hmstring = string.Join("", hm);

            var hour = hmstring.Substring(0,2);
            var minute = hmstring.Substring(2, 2);

            DateTimeOffset MessageTime = new DateTimeOffset(int.Parse(year) + 2000, int.Parse(month), int.Parse(day), int.Parse(hour), int.Parse(minute), 0, TimeSpan.Zero);

            var timestring = string.Join("", MessageTime);

            return (MessageTime.ToString("yyyy/MM/dd H:mm"));

            /*return string.Format("{%y/mm/dd H:mm:ss}", timestring);*/

        }
    }
}
