using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nuget.Org.Npoi.Manager;
using Nuget.Org.Npoi.Manager.Model;

namespace Nuget.Org
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            ExportHelper.Export(new System.Collections.Generic.List<Npoi.Manager.Model.ExportConfigModel>
            {
                new Npoi.Manager.Model.ExportConfigModel{
                    CellStyle = CellStyleEnums.StyleBoldCenter 
                }
            });
        }
    }
}
