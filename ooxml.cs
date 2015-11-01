using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Web;
using System.Data.SqlClient;
using System.Xml;
using System.IO.Packaging;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using NUnit.Framework;

namespace FixPPTLayout
{
    internal class OOXML
    {
        public const string s_sUriWordDoc = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public const string s_sUriContentTypeDoc = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        public const string s_sUriContentTypeSettings = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
        public const string s_sUriMailMergeRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource";
        public const string s_sUriMainPartRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public const string s_sUriRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public const string s_sUriPowerpointUri = "http://schemas.openxmlformats.org/presentationml/2006/main";

        public const string s_sUriSlidePartType = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
        public const string s_sUriSlideRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

        private static void StartElement(XmlTextWriter xw, string sElt)
        {
            xw.WriteStartElement(sElt, s_sUriWordDoc);
        }

        private struct AttrPair
        {
            public string sAttr;
            public string sValue;

            public AttrPair(string sAttrIn, string sValueIn)
            {
                sAttr = sAttrIn;
                sValue = sValueIn;
            }
        };

        static void WriteAttributeString(XmlTextWriter xw, string sAttr, string sValue)
        {
            xw.WriteAttributeString(sAttr, s_sUriWordDoc, sValue);
        }

        private static void EndElement(XmlTextWriter xw)
        {
            xw.WriteEndElement();
        }
        
        private static void WriteElementFull(XmlTextWriter xw, string sElt, string[] rgsVals)
        {
            StartElement(xw, sElt);
            if (rgsVals != null)
                {
                foreach (string s in rgsVals)
                    WriteAttributeString(xw, "val", s);
                }
            EndElement(xw);
        }

        private static void WriteElementFull(XmlTextWriter xw, string sElt, AttrPair[] rgVals)
        {
            StartElement(xw, sElt);
            if (rgVals != null)
                {
                foreach (AttrPair ap in rgVals)
                    WriteAttributeString(xw, ap.sAttr, ap.sValue);
                }
            EndElement(xw);
        }

        public static bool CreateDocSettings(Package pkg, string sDataSource)
        {
            PackagePart prt;
            Stream stm = StmCreatePart(pkg, "/word/settings.xml", s_sUriContentTypeSettings, out prt);

            XmlTextWriter xw = new XmlTextWriter(stm, System.Text.Encoding.UTF8);

            StartElement(xw, "settings");
            WriteElementFull(xw, "view", new[] {"web"});
            StartElement(xw, "mailMerge");
            WriteElementFull(xw, "mainDocumentType", new[] {"email"});
            WriteElementFull(xw, "linkToQuery", (string[])null);
            WriteElementFull(xw, "dataType", new[] {"textFile"});
            WriteElementFull(xw, "connectString", new[] {""});
            WriteElementFull(xw, "query", new[] {String.Format("SELECT * FROM {0}", sDataSource)});
            
            PackageRelationship rel = prt.CreateRelationship( new System.Uri(String.Format("file:///{0}", sDataSource), UriKind.Absolute), TargetMode.External, s_sUriMailMergeRelType);
            StartElement(xw, "dataSource");
            xw.WriteAttributeString("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", rel.Id);
            EndElement(xw);
            StartElement(xw, "odso");
            StartElement(xw, "fieldMapData");
            WriteElementFull(xw, "type", new[] {"dbColumn"});
            WriteElementFull(xw, "name", new[] {"email"});
            WriteElementFull(xw, "mappedName", new[] {"E-mail Address"});
            WriteElementFull(xw, "column", new[] {"0"});
            WriteElementFull(xw, "lid", new[] {"en-US"});
            EndElement(xw); // fieldMapData
            EndElement(xw); // odso
            EndElement(xw); // mailMerge
            EndElement(xw); // settings
            xw.Flush();
            stm.Flush();
            return true;
        }

        static void WritePara(XmlTextWriter xw, string s)
        {
            StartElement(xw, "p");
            StartElement(xw, "r");
            StartElement(xw, "t");
            xw.WriteString(s);
            EndElement(xw); // t
            EndElement(xw); // r
            EndElement(xw); // p
        }

        static void WriteSingleParaCell(XmlTextWriter xw, string sWidth, string sWidthType, string s)
        {
            StartElement(xw, "tc");
            StartElement(xw, "tcPr");
            WriteElementFull(xw, "tcW", new[] {new AttrPair("w", sWidth), new AttrPair("type", sWidthType)});
            EndElement(xw); // tcPr
            WritePara(xw, s);
            EndElement(xw); // tc

        }


        static void StartTable(XmlTextWriter xw, int cCols)
        {
            StartElement(xw, "tbl");
            StartElement(xw, "tblPr");
            WriteElementFull(xw, "tblStyle", new [] { new AttrPair("val", "TableGrid") });
            WriteElementFull(xw, "tblW", new [] { new AttrPair("w", "0"),new AttrPair("type", "auto") });
            WriteElementFull(xw, "tblLook", new [] { new AttrPair("val", "04A0"), new AttrPair("firstRow", "1"), new AttrPair("lastRow", "0"), new AttrPair("firstColumn", "1"), new AttrPair("lastColumn", "0"), new AttrPair("noHBand", "0"), new AttrPair("noVBand", "1") });
            EndElement(xw); // tblPr
            StartElement(xw, "tblGrid");
            while (cCols-- > 0)
                WriteElementFull(xw, "gridCol", new [] { new AttrPair("val", "1440") });
            EndElement(xw); // tblGrid
        }

        static void EndTable(XmlTextWriter xw)
        {
            EndElement(xw); // tbl
        }

#if no

        private static Dictionary<string, string> s_mpSportLevelFriendly = new Dictionary<string, string>()
            {
            {"All Stars SB 9/10's", "SB 10s ALL STARS"},
            {"All Stars SB Majors", "SB Majors ALL STARS"},
            };
 
        static string DescribeGame(GameData.Game gm, int cGames)
        {
            string sDesc;
            string sCount = "";

            if (cGames > 1)
                sCount = String.Format("{0} games, ", cGames);
            else
                sCount = String.Format("{0} game, ", cGames);

            if (gm.TotalSlots - gm.OpenSlots == 0)
                sDesc = "NO UMPIRES";
            else
                sDesc = String.Format("{0} UMPIRE", gm.TotalSlots - gm.OpenSlots);

            return String.Format("{0}{1}", sCount, sDesc);
        }

        static string FriendlySport(GameData.Game gm)
        {
            if (s_mpSportLevelFriendly.ContainsKey(gm.Slots[0].SportLevel))
                return s_mpSportLevelFriendly[gm.Slots[0].SportLevel];

            return gm.Slots[0].SportLevel;
        }

        static void WriteGame(XmlTextWriter xw, GameData.Game gm, int cGames)
        {
            StartElement(xw, "tr");
            WriteSingleParaCell(xw, "0", "auto", gm.Slots[0].Dttm.ToString("M/dd"));
            WriteSingleParaCell(xw, "0", "auto", gm.Slots[0].Dttm.ToString("ddd h tt"));

            WriteSingleParaCell(xw, "0", "auto", FriendlySport(gm));
            WriteSingleParaCell(xw, "0", "auto", gm.Slots[0].SiteShort);

            WriteSingleParaCell(xw, "0", "auto", DescribeGame(gm, cGames));
            EndElement(xw); // tr
        }


        static void AppendGameToSb(GameData.Game gm, int cGames, StringBuilder sb)
        {
            sb.Append(String.Format("<tr><td>{0}<td>{1}", gm.Slots[0].Dttm.ToString("M/dd"), gm.Slots[0].Dttm.ToString("ddd h tt")));
            sb.Append(String.Format("<td>{0}<td>{1}<td class='bold'>{2}", FriendlySport(gm), gm.Slots[0].SiteShort, DescribeGame(gm, cGames)));
        }

        /* C R E A T E  M A I N  D O C */
        /*----------------------------------------------------------------------------
        	%%Function: CreateMainDoc
        	%%Qualified: ArbWeb.OOXML.CreateMainDoc
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static bool CreateMainDoc(Package pkg, string sDataSource, GameData.GameSlots gms, out string sArbiterHelpNeeded)
        {
            PackagePart prt;
            StringBuilder sb = new StringBuilder();

            sb.Append("<div id='D9UrgentHelpNeeded'><h1>HELP NEEDED</h1><h4> The following upcoming games URGENTLY need help! <br>Please <a href=\"https://www.arbitersports.com/Official/SelfAssign.aspx\">SELF ASSIGN</a> now!</h4><style>    table.help td {padding-left: 2mm;padding-right: 2mm;}td.bold {font-weight: bold;}</style> <table class='help' border=1 style='border-collapse: collapse'></div>");

            Stream stm = StmCreatePart(pkg, "/word/document.xml", s_sUriContentTypeDoc, out prt);

            XmlTextWriter xw = new XmlTextWriter(stm, System.Text.Encoding.UTF8);

            StartElement(xw, "document");
            StartElement(xw, "body");
            WritePara(xw, "The following upcoming games URGENTLY need help!");
            WritePara(xw, "If you can work ANY of these games, either sign up on Arbiter, or just reply to this mail and let me know which games you can do. Thanks!");

            StartTable(xw, 5);
            Dictionary<string, List<GameData.Game>> mpSlotGames = new Dictionary<string, List<GameData.Game>>();

            foreach (GameData.Game gm in gms.Games.Values)
                {
                if (gm.TotalSlots - gm.OpenSlots > 1)
                    continue;

                string s = String.Format("{0}-{1}-{2}", gm.Slots[0].Dttm.ToString("yyyyMMdd:HHmm"), gm.Slots[0].SiteShort, gm.TotalSlots - gm.OpenSlots);
                if (!mpSlotGames.ContainsKey(s))
                    mpSlotGames.Add(s, new List<GameData.Game>());

                mpSlotGames[s].Add(gm);
                }

            foreach (List<GameData.Game> plgm in mpSlotGames.Values)
                {
                WriteGame(xw, plgm[0], plgm.Count);
                AppendGameToSb(plgm[0], plgm.Count, sb);
                }
            EndTable(xw);
            sb.Append("</table>");
            EndElement(xw); // body
            EndElement(xw); // document
            
            xw.Flush();
            stm.Flush();

            sArbiterHelpNeeded = sb.ToString();
            return true;
        }
#endif // no
        /* S T M  C R E A T E  P A R T */
        /*----------------------------------------------------------------------------
        	%%Function: StmCreatePart
        	%%Qualified: ArbWeb.OOXML.StmCreatePart
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static Stream StmCreatePart(Package pkg, string sUri, string sContentType, out PackagePart prt)
        {
            Uri uriTeams = new System.Uri(sUri, UriKind.Relative);

            prt = pkg.GetPart(uriTeams);

            List<PackageRelationship> plrel = new List<PackageRelationship>();

            foreach (PackageRelationship rel in prt.GetRelationships())
                {
                plrel.Add(rel);
                }

            prt = null;

            pkg.DeletePart(uriTeams);
            prt = pkg.CreatePart(uriTeams, sContentType);

            foreach (PackageRelationship rel in plrel)
                {
                prt.CreateRelationship(rel.TargetUri, rel.TargetMode, rel.RelationshipType, rel.Id);
                }

            return prt.GetStream(FileMode.Create, FileAccess.Write);
        }


#if no
        /* C R E A T E  M A I L  M E R G E  D O C */
        /*----------------------------------------------------------------------------
        	%%Function: CreateMailMergeDoc
        	%%Qualified: ArbWeb.OOXML.CreateMailMergeDoc
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static void CreateMailMergeDoc(string sTemplate, string sDest, string sDataSource, GameData.GameSlots gms, out string sArbiterHelpNeeded)
        {
            System.IO.File.Copy(sTemplate, sDest);
            Package pkg = Package.Open(sDest, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

            CreateDocSettings(pkg, sDataSource);
            CreateMainDoc(pkg, sDataSource, gms, out sArbiterHelpNeeded);
            pkg.Flush();
            pkg.Close();
        }
#endif

        public static Package PkgOpen(string sPackage, FileMode fm, FileAccess fa, FileShare fs)
        {
            Package pkg = Package.Open(sPackage, fm, fa, fs);

            return pkg;
        }
    }

    public class PresentationX
    {
        private Package m_pkg;

        public PresentationX() { }
        public PresentationX(string sFilename, FileMode fm, FileAccess fa, FileShare fs)
        {
            m_pkg = OOXML.PkgOpen(sFilename, fm, fa, fs);
        }

        public class Slide
        {
            private string m_sName;
            private Uri m_uri;
            private string m_sRelSource; // this is the relId that connects the presentation to the slide
            public string Name => m_sName;
            public Uri Uri => m_uri;
            public string RelSource => m_sRelSource;

            public Slide(string sName, Uri uri, string sRelSource)
            {
                m_sName = sName;
                m_uri = uri;
                m_sRelSource = sRelSource;
            }

            public override string ToString()
            {
                return m_sName;
            }
        }

        public void Close()
        {
            m_pkg.Flush();
            m_pkg.Close();
        }

        PackagePart PrtGetPresentation(out Uri uriPresentationXml)
        {
            PackageRelationshipCollection prc = m_pkg.GetRelationshipsByType(OOXML.s_sUriMainPartRelType);

            if (prc.Count() != 1)
                throw new Exception("must be only 1 main part in a package");

            uriPresentationXml = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), prc.First().TargetUri);
            PackagePart pprt = m_pkg.GetPart(uriPresentationXml);

            return pprt;
        }
        public List<Slide> GetSlides()
        {
            List<Slide> plsld = new List<Slide>();
            // open up presentation.xml and snarf the p:sldId items
            // get the root relationship

            Uri uriPresentationXml;
            PackagePart pprt = PrtGetPresentation(out uriPresentationXml);

            Stream stm = pprt.GetStream(FileMode.Open);

            XmlReaderSettings xrs = new XmlReaderSettings();

            XmlReader xr = XmlReader.Create(stm, xrs);
            
            int nSlideNumberNext = 1;
            while (xr.Read())
                {
                if (xr.IsStartElement("sldId", OOXML.s_sUriPowerpointUri))
                    {
                    string sRelationshipSource = xr.GetAttribute("id", OOXML.s_sUriRelationship);
                    Uri uri = PackUriHelper.ResolvePartUri(uriPresentationXml, pprt.GetRelationship(sRelationshipSource).TargetUri);
                    Slide sld = new Slide(String.Format("Slide{0}", nSlideNumberNext++), uri, sRelationshipSource);

                    plsld.Add(sld);
                    }
                }
            xr.Close();
            stm.Close();
            return plsld;
        }

        private void WriteShallowNode(XmlReader xr, XmlWriter xw)
        {
            if (xr == null)
                {
                throw new ArgumentNullException("xr");
                }
            if (xw == null)
                {
                throw new ArgumentNullException("xw");
                }

            switch (xr.NodeType)
                {
                    case XmlNodeType.Element:
                        xw.WriteStartElement(xr.Prefix, xr.LocalName, xr.NamespaceURI);
                        xw.WriteAttributes(xr, true);
                        if (xr.IsEmptyElement)
                            {
                            xw.WriteEndElement();
                            }
                        break;
                    case XmlNodeType.Text:
                        xw.WriteString(xr.Value);
                        break;
                    case XmlNodeType.Whitespace:
                    case XmlNodeType.SignificantWhitespace:
                        xw.WriteWhitespace(xr.Value);
                        break;
                    case XmlNodeType.CDATA:
                        xw.WriteCData(xr.Value);
                        break;
                    case XmlNodeType.EntityReference:
                        xw.WriteEntityRef(xr.Name);
                        break;
                    case XmlNodeType.XmlDeclaration:
                    case XmlNodeType.ProcessingInstruction:
                        xw.WriteProcessingInstruction(xr.Name, xr.Value);
                        break;
                    case XmlNodeType.DocumentType:
                        xw.WriteDocType(xr.Name, xr.GetAttribute("PUBLIC"), xr.GetAttribute("SYSTEM"),
                                            xr.Value);
                        break;
                    case XmlNodeType.Comment:
                        xw.WriteComment(xr.Value);
                        break;
                    case XmlNodeType.EndElement:
                        xw.WriteFullEndElement();
                        break;
                }
        }

        enum MotionPathPart
        {
            SeekingLine,
            SeekingX,
            ParsingXPre,
            ParsingX,
            SeekingY,
            ParsingYPre,
            ParsingY
        }

        private static void ProcessXYParse(string sPath, int ich, StringBuilder sb, double dConversionFactor, ref int ichStart, ref double x, ref double y, ref MotionPathPart mpp)
        {
            // done!
            double d = Double.Parse(sPath.Substring(ichStart, ich - ichStart));

            if (mpp == MotionPathPart.ParsingX)
                {
                x = d;
                mpp = MotionPathPart.SeekingY;
                ichStart = ich;
                }
            else
                {
                y = d;
                ichStart = ich;
                sb.Append(String.Format("L {0} {1}", x * dConversionFactor, y));
                mpp = MotionPathPart.SeekingLine;
                }

        }

        private static string SFixupMotionPath(string sPath, double dConversionFactor)
        {
            // there's only 1 part that we know how to adjust automatically, the line to
            StringBuilder sb = new StringBuilder();

            int ichStart = 0, ich = 0;
            MotionPathPart mpp = MotionPathPart.SeekingLine;
            double x = 0.0, y = 0.0;

            while (ich < sPath.Length)
                {
                if (mpp == MotionPathPart.SeekingLine)
                    {
                    if (sPath[ich] == 'L')
                        {
                        if (ich - ichStart > 0)
                            sb.Append(sPath.Substring(ichStart, ich - ichStart));
                        ichStart = ++ich;
                        mpp = MotionPathPart.SeekingX;
                        continue;
                        }
                    ich++;
                    continue;
                    }
                if (mpp == MotionPathPart.SeekingX || mpp == MotionPathPart.SeekingY)
                    {
                    if (Char.IsWhiteSpace(sPath[ich]))
                        {
                        ichStart = ++ich;
                        continue;
                        }
                    mpp = mpp == MotionPathPart.SeekingX ? MotionPathPart.ParsingXPre : MotionPathPart.ParsingYPre;
                    }

                if (mpp == MotionPathPart.ParsingXPre || mpp == MotionPathPart.ParsingYPre)
                    {
                    if (sPath[ich] == '-')
                        {
                        ich++;
                        }
                    if (Char.IsDigit(sPath[ich]) || sPath[ich] == '.')
                        {
                        mpp = mpp == MotionPathPart.ParsingXPre ? MotionPathPart.ParsingX : MotionPathPart.ParsingY;
                        // FALLTHROUGH
                        }
                    else
                        {
                        // anything else here is a corrput value
                        throw new Exception(string.Format("Motion string illegal: {0}", sPath));
                        }
                    }

                if (mpp == MotionPathPart.ParsingX || mpp == MotionPathPart.ParsingY)
                    {
                    if (Char.IsDigit(sPath[ich]) || sPath[ich] == 'E' || sPath[ich] == '-' || sPath[ich] == '.')
                        {
                        ich++;
                        continue;
                        }
                    if (Char.IsWhiteSpace(sPath[ich]))
                        {
                        ProcessXYParse(sPath, ich, sb, dConversionFactor, ref ichStart, ref x, ref y, ref mpp);
                        continue;
                        }
                    }
                throw new Exception("shouldn't get here");
                }

            if (mpp == MotionPathPart.ParsingY)
                ProcessXYParse(sPath, ich, sb, dConversionFactor, ref ichStart, ref x, ref y, ref mpp);

            if (ichStart < ich)
                sb.Append(sPath.Substring(ichStart, ich - ichStart));
            return sb.ToString();
        }

        [TestCase("M 0 0", 1.0, "M 0 0")]
        [TestCase("L 0 0", 1.0, "L 0 0")]
        [TestCase("M 0 0 L 0 0", 1.0, "M 0 0 L 0 0")]
        [TestCase("M 0 0 L 0 0 U 0 0", 1.0, "M 0 0 L 0 0 U 0 0")]
        [TestCase("L 0.0 0.0", 1.0, "L 0 0")]
        [TestCase("L 0.1 1.1", 1.0, "L 0.1 1.1")]
        [TestCase("L .1 1.1", 1.0, "L 0.1 1.1")]
        [TestCase("L 1 1.1", 1.0, "L 1 1.1")]
        [TestCase("L 1E-1 1.1", 1.0, "L 0.1 1.1")]
        [TestCase("L 1.97131E-6 1.1", 1.0, "L 1.97131E-06 1.1")]
        [TestCase(" L -0.18341 0", 0.75, " L -0.1375575 0")]
        [Test]
        public static void TestFixSlideAnimMotion(string sPath, double dConversionFactor, string sExpected)
        {
            Assert.AreEqual(sExpected, SFixupMotionPath(sPath, dConversionFactor));
        }

        public void FixSlideAnimMotion(Slide sld, Uri uri)
        {
            string sSlideName = sld.Name;
            Uri uriNew = new Uri(String.Format("\\fixed\\{0}.xml", sSlideName), UriKind.Relative);
            Uri uriPartNew = PackUriHelper.CreatePartUri(uriNew);

            PackagePart pprtNew = m_pkg.CreatePart(uriPartNew, OOXML.s_sUriSlidePartType);
            Stream stm = pprtNew.GetStream(FileMode.Create, FileAccess.Write);

            PackagePart pprtSource = m_pkg.GetPart(uri);
            Stream stmSource = pprtSource.GetStream(FileMode.Open, FileAccess.Read);
            
            XmlReaderSettings xrs = new XmlReaderSettings();
            XmlReader xr = XmlReader.Create(stmSource, xrs);
            XmlWriter xw = XmlWriter.Create(stm);

            while (xr.Read())
                {
                if (xr.IsStartElement("animMotion", OOXML.s_sUriPowerpointUri) &&
                    xr.GetAttribute("origin") == "layout")
                    {
                    string sPath = xr.GetAttribute("path");

                    sPath = SFixupMotionPath(sPath, 0.75);
                    xw.WriteStartElement(xr.Prefix, xr.LocalName, xr.NamespaceURI);
                    while (xr.MoveToNextAttribute())
                        {
                        if (xr.LocalName == "path")
                            xw.WriteAttributeString("path", sPath);
                        else
                            xw.WriteAttributeString(xr.LocalName, xr.NamespaceURI, xr.Value);
                        }
                    if (xr.IsEmptyElement)
                        xw.WriteEndElement();
                    }
                else
                    WriteShallowNode(xr, xw);
                }
            xr.Close();
            xw.Flush();
            xw.Close();
            stm.Close();
            stmSource.Close();

            // at this point, we have a new part, uriPartNew / pprtNew, but we have to repoint
            // the old relationships to this one...
            Uri uriPresentationXml;
            PackagePart prt = PrtGetPresentation(out uriPresentationXml);
            prt = null;

            FixupRelationship(uriPresentationXml, uri, uriPartNew);

//            prt.DeleteRelationship(sld.RelSource);
            //prt.CreateRelationship(PackUriHelper.ResolvePartUri(uriPresentationXml, uriPartNew), TargetMode.Internal, OOXML.s_sUriSlideRelType, sld.RelSource);

            // now copy the relationships
            foreach (PackageRelationship pr in pprtSource.GetRelationships())
                {
                Uri uriNewPr;

                if (pr.TargetMode == TargetMode.External)
                    {
                    uriNewPr = pr.TargetUri;
                    }
                else
                    {
                    Uri uriFull = PackUriHelper.ResolvePartUri(pprtSource.Uri, pr.TargetUri);
                    uriNewPr = PackUriHelper.ResolvePartUri(uriPartNew, uriFull);

                    // if its a notes slide, then we have to fix up its back pointing relationships
                    FixupRelationship(uriFull, uri, uriPartNew);
                    }

                pprtNew.CreateRelationship(uriNewPr, pr.TargetMode, pr.RelationshipType, pr.Id);
                }
        }

        void FixupRelationship(Uri uriReferringPart, Uri uriOldTarget, Uri uriPartNew)
        {
            PackagePart prt = m_pkg.GetPart(uriReferringPart);
            List<string> plsIds = new List<string>();

            foreach (PackageRelationship pr in prt.GetRelationships())
                {
                if (pr.TargetMode == TargetMode.External)
                    continue;

                Uri uriPr = PackUriHelper.ResolvePartUri(uriReferringPart, pr.TargetUri);
                if (uriPr.Equals(uriOldTarget))
                    {
                    plsIds.Add(pr.Id);
                    }
                }

            foreach (string s in plsIds)
                {
                prt.DeleteRelationship(s);
                prt.CreateRelationship(PackUriHelper.ResolvePartUri(uriReferringPart, uriPartNew), TargetMode.Internal,
                                       OOXML.s_sUriSlideRelType, s);
                }
        }
    }
}