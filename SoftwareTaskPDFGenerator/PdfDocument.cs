using System.Collections.Generic;

namespace SoftwareTaskPDFGenerator
{
    internal class PdfDocument
    {
        private List<Employees> employees;

        public PdfDocument(List<Employees> employees)
        {
            this.employees = employees;
        }
    }
}