using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Text;

namespace match_employee.Models
{
    public class Employee
    {
        [Key]
        public string Number { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Gender { get; set; }
        public DateTime DateOfBirth { get; set; }
        

        [NotMapped]
        public int? ExcelRow { get; set; }
        [NotMapped]
        public string? NumberOfEmployees { get; set; }

        // Use the invariant culture (non-case-insensitive) and single character misspelling to compare strings
        private bool CompareStrings(string str1, string str2)
        {
            string normalizedStr1 = NormalizeString(str1);
            string normalizedStr2 = NormalizeString(str2);

            // Check if the two strings are equal
            if (normalizedStr1.Equals(normalizedStr2, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // Check for a single character misspelling
            return IsOneCharDifferent(normalizedStr1, normalizedStr2);
        }

        // Normalize the string to remove accents
        private string NormalizeString(string str)
        {
            var normalized = str.Normalize(NormalizationForm.FormD);
            char[] array = new char[normalized.Length];
            int index = 0;

            foreach (char c in normalized)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                {
                    array[index++] = char.ToLower(c); // Convert to lower case
                }
            }
            return new string(array, 0, index); // Return normalized string
        }

        //Compare two string with ignore of single letter misspelling
        private bool IsOneCharDifferent(string str1, string str2)
        {
            // Check length difference
            if (Math.Abs(str1.Length - str2.Length) > 1)
            {
                return false; // Lengths differ by more than 1, cannot be similar
            }

            int mismatchCount = 0;
            int length1 = str1.Length;
            int length2 = str2.Length;

            // Use two pointers to compare characters
            int i = 0;
            int j = 0;

            while (i < length1 && j < length2)
            {
                if (str1[i] != str2[j])
                {
                    mismatchCount++;
                    if (mismatchCount > 1)
                    {
                        return false; // More than one mismatch
                    }

                    // If lengths are the same, move both pointers
                    if (length1 == length2)
                    {
                        i++;
                        j++;
                    }
                    else if (length1 > length2)
                    {
                        // If str1 is longer, move only the str1 pointer
                        i++;
                    }
                    else
                    {
                        // If str2 is longer, move only the str2 pointer
                        j++;
                    }
                }
                else
                {
                    // Characters match, move both pointers
                    i++;
                    j++;
                }
            }

            // If we exit the loop, we need to check if there's one extra character at the end
            if (i < length1 || j < length2)
            {
                mismatchCount++;
            }

            return mismatchCount <= 1;
        }
        // Compare two Employee model
        public bool CompareEmployee(Employee employee)
        {
            return Gender.Equals(employee.Gender, StringComparison.OrdinalIgnoreCase)
                && DateOfBirth == employee.DateOfBirth;
        }
    }
}
