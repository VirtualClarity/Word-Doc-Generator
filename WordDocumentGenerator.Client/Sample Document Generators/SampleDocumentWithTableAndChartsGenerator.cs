// ----------------------------------------------------------------------
// <copyright file="SampleDocumentWithTableAndChartsGenerator.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Client
{
    using System.Linq;
    using DocumentFormat.OpenXml.Drawing.Charts;
    using DocumentFormat.OpenXml.Packaging;
    using WordDocumentGenerator.Library;

    /// <summary>
    /// Document generator for template having table and charts
    /// </summary>
    public class SampleDocumentWithTableAndChartsGenerator : SampleDocumentWithTableGenerator
    {
        private readonly LineChartData lineChartData;
        private readonly ScatterChartData scatterChartData;

        /// <summary>
        /// Initializes a new instance of the <see cref="SampleDocumentWithTableAndChartsGenerator"/> class.
        /// </summary>
        /// <param name="generationInfo">The generation info.</param>
        /// <param name="lineChartData">The line chart data.</param>
        /// <param name="scatterChartData">The scatter chart data.</param>
        public SampleDocumentWithTableAndChartsGenerator(DocumentGenerationInfo generationInfo, LineChartData lineChartData, ScatterChartData scatterChartData)
            : base(generationInfo)
        {
            this.lineChartData = lineChartData;
            this.scatterChartData = scatterChartData;
        }

        /// <summary>
        /// Refreshes the charts.
        /// </summary>
        /// <param name="mainDocumentPart">The main document part.</param>
        protected override void RefreshCharts(MainDocumentPart mainDocumentPart)
        {
            if (mainDocumentPart != null)
            {
                foreach (ChartPart chartPart in mainDocumentPart.ChartParts)
                {
                    Chart chart = chartPart.ChartSpace.Elements<Chart>().FirstOrDefault();

                    if (chart != null)
                    {
                        DocumentFormat.OpenXml.Drawing.Charts.ScatterChart scatterChart = chart.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ScatterChart>().FirstOrDefault();
                        DocumentFormat.OpenXml.Drawing.Charts.Line3DChart lineChart = chart.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Line3DChart>().FirstOrDefault();

                        if (scatterChart != null)
                        {                            
                            ScatterChartEx chartEx = new ScatterChartEx(chartPart, this.scatterChartData);
                            chartEx.Refresh();
                        }

                        if (lineChart != null)
                        {                            
                            Line3DChartEx chartEx = new Line3DChartEx(chartPart, this.lineChartData);
                            chartEx.Refresh();
                        }
                    }

                    chartPart.ChartSpace.Save();
                }
            }
        }
    }
}
