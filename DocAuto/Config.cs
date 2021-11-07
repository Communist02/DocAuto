using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace DocAuto
{
    class Config
    {
        public string[] lastDocument { get; set; }

        public Config()
        {
            lastDocument = Array.Empty<string>();
        }

        public void addDocument(string filePath)
        {
            lastDocument = (new string[] { filePath }).Concat(lastDocument).Distinct().ToArray();
        }

        public void LastDocumentClear()
        {
            lastDocument = Array.Empty<string>();
        }

        public int CountLastDocument()
        {
            return lastDocument.Length;
        }
    }
}
