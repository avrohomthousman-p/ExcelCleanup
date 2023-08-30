using System;

namespace CompatableExcelCleaner
{
    /// <summary>
    /// Stores the information identifying a single worksheet - the report name and worksheet number.
    /// </summary>
    internal class Worksheet
    {
        public int worksheetNum { get; }
        public string reportName { get; }


        public Worksheet(string reportName, int worksheetNumber)
        {
            if (reportName == null)
            {
                throw new ArgumentNullException("Cannot create worksheet for null report");
            }
            if (worksheetNumber < 0)
            {
                throw new ArgumentException("worksheet number cannot be negative. Got " + worksheetNumber);
            }


            this.worksheetNum = worksheetNumber;
            this.reportName = reportName;
        }



        public override int GetHashCode()
        {
            int hashcode = 13;

            hashcode = hashcode * 23 + worksheetNum;
            hashcode = hashcode * 23 + reportName.GetHashCode();

            return hashcode;
        }



        public override bool Equals(Object other)
        {
            Worksheet cast = other as Worksheet;
            if ((object)cast == null)
            {
                return false;
            }


            return this == cast;
        }


        public static bool operator ==(Worksheet one, Worksheet two)
        {
            // If both are null, or both are same instance, return true.
            if (System.Object.ReferenceEquals(one, two))
            {
                return true;
            }

            // If one is null, but not both, return false.
            if (((object)one == null) || ((object)one == null))
            {
                return false;
            }


            return one.worksheetNum == two.worksheetNum && one.reportName == two.reportName;
        }


        public static bool operator !=(Worksheet one, Worksheet two)
        {
            return !(one == two);
        }
    }
}
