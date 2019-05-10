using EnvDTE;
using EnvDTE80;
using System.Collections.Generic;
using System.Linq;
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
        string sEmbeddedResource = "EmbeddedResource";
        string sAssembly = "assembly";
        string sName = "name";
        string sValue = "value";

        public void UpdateSolution(DTE2 dte)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            if (dte != null && dte.Solution != null)
                foreach (Project proj in dte.Solution.Projects)
                {
                    string projFileName = proj.FileName;
                    if (string.IsNullOrWhiteSpace(projFileName))
                        continue;
                    UpdateProj(projFileName);
                }
        }

        void UpdateProj(string projFileName)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(projFileName);

            if (UpdateProjCore(doc, true) | UpdateProjCore(doc, false, GetParentDirectory(projFileName)))
                doc.Save(projFileName);
        }

        private bool UpdateProjCore(XmlDocument doc, bool references, string docParentDirectory = "")
        {
            string tagName = references ? sReference : sEmbeddedResource;
            List<XmlNode> projNodes = doc.GetElementsByTagName(tagName).Cast<XmlNode>().ToList();
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
                                    DeleteLicx(GetFilePath(docParentDirectory, nodeValue), node);
                                    res = true;
                                }
                                else if (nodeValue.Contains(sResx))
                                    UpdateResx(GetFilePath(docParentDirectory, nodeValue));
                            }
                    }
            }
            return res;
        }

        void DeleteLicx(string filePath, XmlNode node)
        {
            node.ParentNode.RemoveChild(node);
            if (File.Exists(filePath))
                File.Delete(filePath);
        }

        string GetFilePath(string docParentDirectory, string nodeValue)
        {
            return docParentDirectory + string.Format(@"\{0}", nodeValue);
        }

        void UpdateResx(string filePath)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(filePath);

            XmlNodeList assemblyNodes = doc.GetElementsByTagName(sAssembly);
            XmlNodeList valueNodes  = doc.GetElementsByTagName(sValue);

            if (UpdateAssemblyNodes(assemblyNodes) | UpdateValueNodes(valueNodes))
                doc.Save(filePath);
        }

        bool UpdateValueNodes(XmlNodeList projNodes)
        {
            string nodeValue;
            string newNodeValue;
            XmlNode node;
            bool res = false;
            for (int nodeIndex = 0; nodeIndex < projNodes.Count; nodeIndex++)
            {
                node = projNodes[nodeIndex];
                nodeValue = node.InnerText;
                newNodeValue = RemoveReferenceVersion(nodeValue);
                if(nodeValue != newNodeValue)
                {
                    node.InnerText = newNodeValue;
                    res = true;
                }
            }
            return res;
        }

        bool UpdateAssemblyNodes(XmlNodeList projNodes)
        {
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
            return res;
        }

        string RemoveReferenceVersion(string str)
        {
            int removeIndex = str.LastIndexOf(sDevExpress);
            if (removeIndex >= 0)
                str = RemoveReferenceVersionCore(str, removeIndex);
            return str;
        }

        string RemoveReferenceVersionCore(string str, int startIndex)
        {
            int removeIndex = str.IndexOf(',', startIndex);
            if (removeIndex >= 0)
                str = str.Remove(removeIndex);
            return str;
        }

        string GetParentDirectory(string projFileName)
        {
            return Directory.GetParent(projFileName).FullName;
        }
    }
}
