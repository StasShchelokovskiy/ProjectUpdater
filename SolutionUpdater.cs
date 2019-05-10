using EnvDTE;
using EnvDTE80;
using System.IO;
using System.Xml;

namespace ProjectUpdater
{
    public class SolutionUpdater
    {
        string sItemGroup = "ItemGroup";
        string sReference = "Reference";
        string sInclude = "Include";
        string sDevExpress = "DevExpress";
        string sLicx = "licx";
        string sResx = "resx";
        string sProperties = @"\Properties";
        string sMyProject = @"\My Project";
        string sEmbeddedResource = "EmbeddedResource";
        string sAssembly = "assembly";
        string sName = "name";

        public void UpdateSolution(DTE2 dte)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            if (dte != null && dte.Solution != null)
                foreach (Project proj in dte.Solution.Projects)
                {
                    string projFileName = proj.FileName;
                    if (string.IsNullOrWhiteSpace(projFileName))
                        continue;
                    DeleteLicx(projFileName);
                    UpdateProj(projFileName);
                }
        }

        void DeleteLicx(string projFileName)
        {
            string propertiesFolder = GetParentDirectory(projFileName);
            propertiesFolder += Directory.Exists(propertiesFolder + sProperties) ? sProperties : sMyProject;
            string[] allFiles = Directory.GetFiles(propertiesFolder, string.Format("*.{0}", sLicx));
            foreach (string file in allFiles)
                File.Delete(file);
        }

        void UpdateProj(string projFileName)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(projFileName);

            if (PerformUpdate(doc, true) | PerformUpdate(doc, false, GetParentDirectory(projFileName)))
                doc.Save(projFileName);
        }

        private bool PerformUpdate(XmlDocument doc, bool references, string docParentDirectory = "")
        {
            string tagName = references ? sReference : sEmbeddedResource;
            XmlNodeList projNodes = doc.GetElementsByTagName(tagName);
            string nodeValue;
            string newNodeValue;
            XmlNode node;
            bool res = false;
            for (int nodeIndex = 0; nodeIndex < projNodes.Count; nodeIndex++)
            {
                node = projNodes[nodeIndex];
                if (node.ParentNode.Name == sItemGroup)
                    for (int attributeIndex = 0; attributeIndex < node.Attributes.Count; attributeIndex++)
                    {
                        nodeValue = node.Attributes[attributeIndex].Value;
                        if (node.Attributes[attributeIndex].Name == sInclude)
                            if (references)
                            {
                                newNodeValue = RemoveReferenceVersion(nodeValue);
                                if (nodeValue != newNodeValue)
                                {
                                    node.Attributes[attributeIndex].Value = newNodeValue;
                                    res = true;
                                }
                            }
                            else
                            {
                                if (nodeValue.Contains(sLicx))
                                {
                                    node.ParentNode.RemoveChild(node);
                                    res = true;
                                }
                                else if (nodeValue.Contains(sResx))
                                    ProcessResx(docParentDirectory + string.Format(@"\{0}", nodeValue));
                            }
                    }
            }
            return res;
        }

        void ProcessResx(string filePath)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(filePath);
            XmlNodeList projNodes = doc.GetElementsByTagName(sAssembly);
            string nodeValue;
            string newNodeValue;
            XmlNode node;
            bool res = false;
            for (int nodeIndex = 0; nodeIndex < projNodes.Count; nodeIndex++)
            {
                node = projNodes[nodeIndex];
                for (int attributeIndex = 0; attributeIndex < node.Attributes.Count; attributeIndex++)
                {
                    nodeValue = node.Attributes[attributeIndex].Value;
                    if (node.Attributes[attributeIndex].Name == sName)
                    {
                        newNodeValue = RemoveReferenceVersion(nodeValue);
                        if (nodeValue != newNodeValue)
                        {
                            node.Attributes[attributeIndex].Value = newNodeValue;
                            res = true;
                        }
                    }
                }
            }
            if (res)
                doc.Save(filePath);
        }

        private string RemoveReferenceVersion(string str)
        {
            if (str.Contains(sDevExpress))
            {
                int startIndex = str.IndexOf(',');
                if (startIndex >= 0)
                    str = str.Remove(startIndex);
            }
            return str;
        }

        private string GetParentDirectory(string projFileName)
        {
            return Directory.GetParent(projFileName).FullName;
        }
    }
}
