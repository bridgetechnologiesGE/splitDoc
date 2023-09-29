using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Markup;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SplitPdf
{
    internal static class Program
    {
        internal enum SplitError : int
        {
            Ok = 0,
            Unk = -1,
            InvalidArgs = 1,
            ConfNotFound = 2,
            InvalidConf = 3,
            SplitMaskMissing = 4,
            InputError = 5,
            ConfDocMissingInput = 6,
            ConfDocInputNotFound = 7,
            ConfDocInvalidInvRange = 8,
            ConfDocInvalidDepthRange = 9,
        }

        internal class SplitException : Exception
        {
            public SplitError SplitError { get; private set; }

            public SplitException(SplitError splitError)
            {
                this.SplitError = splitError;
            }
        }

        public static JObject JsonFromFile(this string inPath)
        {
            using (var rdr = File.OpenText(inPath))
            using (var tRdr = new JsonTextReader(rdr))
                return new JsonSerializer().Deserialize<JObject>(tRdr);
        }

        public static bool AddReader(this Dictionary<string, PdfReader> inDict, string inPath)
        {
            try
            {
                if (!inDict.ContainsKey(inPath))
                    inDict.Add(inPath, new PdfReader(inPath));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string metanize(this string mask, Dictionary<string, string> metadata)
        {
            IEnumerable<string> tokenize()
            {
                foreach (var tkn in Regex.Split(mask, "(%%%.+?%%%)"))
                {
                    switch (tkn)
                    {
                        case var _ when string.IsNullOrWhiteSpace(tkn):
                            break;

                        case var _ when !tkn.StartsWith("%%%"):
                        case var _ when !tkn.EndsWith("%%%"):
                        case var _ when tkn.Length <= 6:
                            yield return tkn;
                            break;

                        default:
                            var key = tkn.Substring(3, tkn.Length - 6);
                            yield return metadata.ContainsKey(key) ? metadata[key] : key;
                            break;
                    }
                }
            }

            return string.Join("", tokenize());
        }

        public static IEnumerable<T> JsonArray<T>(this JObject item, string ele) where T : JToken
            => (item.Property(ele).Value as JArray).Children<T>();

        static int Main(string[] args)
        {
            try
            {
                if (args.Length != 1)
                    throw new SplitException(SplitError.InvalidArgs);

                var conf = default(JObject);

                try
                {
                    conf = args[0].JsonFromFile();
                }
                catch (FileNotFoundException)
                {
                    throw new SplitException(SplitError.ConfNotFound);
                }
                catch
                {
                    throw new SplitException(SplitError.InvalidConf);
                }

                var inFiles = new Dictionary<string, PdfReader>();
                var outInvMask = conf.Value<string>("splitInvName");
                var outDepthMask = conf.Value<string>("splitDepthName");
                var outFolder = conf.Value<string>("outFolder");

                if (string.IsNullOrWhiteSpace(outInvMask) && string.IsNullOrWhiteSpace(outDepthMask))
                    throw new SplitException(SplitError.SplitMaskMissing);

                if (string.IsNullOrWhiteSpace(outFolder))
                    outFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "out");

                if (!Directory.Exists(outFolder))
                    Directory.CreateDirectory(outFolder);

                var _sc = StringComparer.InvariantCultureIgnoreCase;
                var inDict = new Dictionary<string, PdfReader>(_sc);

                foreach (var ele in conf.JsonArray<JObject>("splitDocs"))
                {
                    var inPath = ele.Value<string>("inputFile");
                    var invPages = new[] { ele.Value<int>("fromInvoice"), ele.Value<int>("toInvoice") };
                    var depthPages = new[] { ele.Value<int>("fromDepth"), ele.Value<int>("toDepth") };

                    foreach (var check in new (Func<bool> isOk, SplitError err)[]
                    {
                        ( isOk : () => !string.IsNullOrWhiteSpace(inPath), err : SplitError.ConfDocMissingInput ),
                        ( isOk : () => File.Exists(inPath), err : SplitError.ConfDocInputNotFound ),
                        ( isOk : () => inDict.AddReader(inPath), err: SplitError.InputError ),
                        ( isOk : () => invPages.All(ix => ix > 0) && invPages[1] >= invPages[0] &&
                        invPages[1] <= inDict[inPath].NumberOfPages, err : SplitError.ConfDocInvalidInvRange ),
                        ( isOk : () => depthPages.All(ix => ix > 0) && depthPages[1] >= depthPages[0] &&
                        depthPages[1] <= inDict[inPath].NumberOfPages, err : SplitError.ConfDocInvalidDepthRange ),
                    })
                    {
                        if (check.isOk()) continue;
                        throw new SplitException(check.err);
                    }

                    var metaDepth = new Dictionary<string, string>(_sc);

                    foreach (var property in ele.Value<JObject>("metakeys").Properties())
                        metaDepth.Add(property.Name, property.Value.ToString());

                    foreach (var entry in new (string mask, int[] range)[]
                    {
                        (mask: outInvMask, range: invPages),
                        (mask: outDepthMask, range: depthPages)
                    })
                    {
                        if (string.IsNullOrWhiteSpace(entry.mask))
                            continue;

                        var outFullPath = Path.Combine(outFolder,
                            entry.mask.metanize(metaDepth));

                        using (var str = File.Create(outFullPath))
                        {
                            var doc = new Document(PageSize.A4);
                            var copier = new PdfCopy(doc, str);
                            var numOfPages = entry.range[1] - entry.range[0] + 1;

                            doc.Open();
                            foreach (var pageId in Enumerable.Range(entry.range[0], numOfPages))
                                copier.AddPage(copier.GetImportedPage(inDict[inPath], pageId));
                            copier.Flush();
                            doc.Close();
                        }
                    }
                }

                return (int)SplitError.Ok;
            }
            catch (Exception ex)
            {
                return (int)((ex as SplitException)?.SplitError ?? SplitError.Unk);
            }
        }
    }
}
