using CsvHelper;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace SoftwareTaskPDFGenerator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Set the license type to Community
            QuestPDF.Settings.License = LicenseType.Community;

            // Read the CSV file and convert it to a list of employees
            string csvFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resources", "sample - data.csv");  // CSV file path
            var employees = ReadEmployeesFromCsv(csvFilePath);
            
            //Main PDF Generation
            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.DefaultTextStyle(style => style.FontFamily("Open Sans"));
                    page.Margin(50);
                    page.Content().Column(column =>
                    {
                        //Iterate through each employee
                        foreach (var employee in employees)
                        {
                            
                            //Top of the page
                            column.Item().Row(row =>
                            {
                                // Left side: Employee's full name, Arabic name, and job title
                                row.RelativeColumn().Stack(stack =>
                                {
                                    stack.Item().PaddingBottom(50).Text(employee.FullName).FontSize(36).Bold();
                                    stack.Item().PaddingBottom(10).AlignCenter().Text(employee.ArabicName).FontSize(18).FontFamily("Arial Unicode MS");
                                    stack.Item().PaddingBottom(10).Text(employee.JobTitle).FontSize(14).Bold();
                                    stack.Item().PaddingBottom(10).Text(text =>
                                    {
                                        text.Span("Join Date: ").Bold().FontSize(10);
                                        text.Span(employee.JoiningDate).FontSize(10);
                                    });

                                });

                                // Right side: Employee's work information
                                row.RelativeColumn().Stack(stack =>
                                {
                                    stack.Item().PaddingVertical(5).LineHorizontal(5).LineColor(Colors.Black);
                                    stack.Item().PaddingBottom(5).Text("Work Information").FontSize(14).Bold().FontColor("#397e89");
                                    stack.Item().PaddingBottom(5).Text($"ID: {employee.EmployeeID}").FontSize(14).Bold().FontColor("#397e89");
                                    stack.Item().PaddingBottom(5).Text(text =>
                                    {
                                        text.Span("Company: ").Bold().FontSize(10);
                                        text.Span(employee.Company).FontSize(10);
                                    });

                                    stack.Item().PaddingBottom(5).Text(text =>
                                    {
                                        text.Span("Department: ").Bold().FontSize(10);
                                        text.Span(employee.Department).FontSize(10);
                                    });
                                    stack.Item().PaddingBottom(5).Text(text =>
                                    {
                                        text.Span("Branch: ").Bold().FontSize(10);
                                        text.Span(employee.Branch).FontSize(10);
                                    });
                                    stack.Item().PaddingBottom(5).Text(text =>
                                    {
                                        text.Span("Mobile: ").Bold().FontSize(10);
                                        text.Span(employee.WorkPhoneNumber).FontSize(10);
                                    });
                                    stack.Item().PaddingBottom(5).Text(text =>
                                    {
                                        text.Span("Email: ").Bold().FontSize(10);
                                        text.Span(employee.WorkEmailAddress).FontSize(10);
                                    });
                                    stack.Item().PaddingBottom(5).Text(text =>
                                    {
                                        text.Span("Direct Manager: ").Bold().FontSize(10);
                                        text.Span(employee.DirectManager).FontSize(10);
                                    });
                                    stack.Item().PaddingBottom(5).Text(text =>
                                    {
                                        text.Span("Supervisor: ").Bold().FontSize(10);
                                        text.Span(employee.Supervisor).FontSize(10);
                                    });
                                });
                            });

                            column.Item().PaddingVertical(2).LineHorizontal(5).LineColor(Colors.Black);

                            //Body of the page
                            column.Item().Row(row =>
                            {
                                // Left side: Personal information
                                row.RelativeColumn(1).Stack(stack =>
                                {
                                    stack.Item().Text("Personal\nInformation").FontColor("#397e89").FontSize(14).Bold();  
                                });

                                // Personal Information Section Middle Column
                                row.RelativeColumn(2).AlignCenter().Stack(stack =>
                                {
                                    stack.Item().Padding(6).Text(text =>
                                    {
                                        text.Span("Citizenship: ").Bold().FontSize(10);
                                        text.Span(employee.Citizenship).FontSize(10);
                                    });

                                    stack.Item().Padding(6).Text(text =>
                                    {
                                        text.Span("Country of Birth: ").Bold().FontSize(10);
                                        text.Span(employee.CountryOfBirth).FontSize(10);
                                    });

                                    stack.Item().Padding(6).Text(text =>
                                    {
                                        text.Span("Place of Birth: ").Bold().FontSize(10);
                                        text.Span(employee.PlaceOfBirth).FontSize(10);
                                    });

                                    stack.Item().Padding(6).Text(text =>
                                    {
                                        text.Span("Blood Group: ").Bold().FontSize(10);
                                        text.Span(employee.BloodGroup).FontSize(10);
                                    });

                                    stack.Item().Padding(6).Text(text =>
                                    {
                                        text.Span("Birthday: ").Bold().FontSize(10);
                                        text.Span(employee.Birthdate).FontSize(10);
                                        text.Span("  ");
                                        text.Span(employee.Gender).FontSize(10);
                                        text.Span("  ");
                                        text.Span(employee.Religion).FontSize(10);
                                    });
                                    stack.Item().Padding(6).Text(text =>
                                    {
                                        text.Span("Marital Status: ").Bold().FontSize(10);
                                        text.Span(employee.MaritalStatus).FontSize(10);
                                        text.Span("  ");
                                        text.Span(employee.SpouseName).FontSize(10);
                                        text.Span("  ");
                                        text.Span(employee.SpouseBirthday).FontSize(10);
                                    });

                                   
                                });

                                // Right side: Employee's image
                                row.ConstantColumn(100).Element(imageContainer =>
                                {
                                    var imagePath = GetImagePath(employee.EmployeeID);
                                    if (File.Exists(imagePath))
                                    {
                                        
                                        imageContainer.Image(imagePath);
                                    }
                                    else
                                    {
                                        imageContainer.Text("No Image Available").FontSize(12);
                                    }
                                });

                            });

                            //Identification Section
                            column.Item().Row(row =>
                            {
                                row.RelativeColumn(1).Stack(stack =>
                                {
                                    stack.Item().Text("Identification").FontColor("#397e89").FontSize(14).Bold();
                                });

                                // Identification Section Middle Column
                                row.RelativeColumn(2).PaddingLeft(2).Stack(stack =>
                                {
                                    stack.Item().Padding(6).AlignLeft().Text(text =>
                                    {
                                        text.Span("National ID: ").Bold().FontSize(10);
                                        text.Span(employee.NationalID).FontSize(10);
                                    });

                                    stack.Item().Padding(6).PaddingBottom(10).AlignLeft().Text(text =>
                                    {
                                        text.Span("Passport No.: ").Bold().FontSize(10);
                                        text.Span(employee.PassportNumber).FontSize(10);
                                    });
                                });
                            });

                            //Personal Contact Section
                            column.Item().Row(row =>
                            {
                                row.RelativeColumn(1).PaddingLeft(2).Stack(stack =>
                                {
                                    stack.Item().Text("Personal Contact").FontColor("#397e89").FontSize(14).Bold();
                                });
                                
                                // Personal Contact Section Middle Column
                                row.RelativeColumn(2).Stack(stack =>
                                {
                                    stack.Item().Padding(6).AlignLeft().Text(text =>
                                    {
                                        text.Span("Address: ").Bold().FontSize(10);
                                        text.Span(employee.HomeAddress).FontSize(10);
                                    });

                                    stack.Item().Padding(6).AlignLeft().Text(text =>
                                    {
                                        text.Span("Mobile: ").Bold().FontSize(10);
                                        text.Span(employee.PersonalPhoneNumber).FontSize(10);
                                    });

                                    stack.Item().Padding(6).AlignLeft().Text(text =>
                                    {
                                        text.Span("Email: ").Bold().FontSize(10);
                                        text.Span(employee.PersonalEmailAddress).FontSize(10);
                                    });
                                });
                            });

                            //Bottom of the Page                   
                            column.Item().Row(row =>
                            {
                                //Education Column
                                row.RelativeColumn().Stack(stack =>
                                {
                                    stack.Item().Padding(2).LineHorizontal(5).LineColor(Colors.Black);
                                    stack.Item().Text("Education").FontColor("#397e89").FontSize(12).Bold();
                                    stack.Item().Padding(3).Text(employee.Education).FontSize(9);
                                    stack.Item().Padding(3).Text(employee.School).FontSize(9);
                                    stack.Item().Padding(3).Text(employee.FieldOfStudy).FontSize(9);
                                });

                                // Emergency Contact 1 Column
                                row.RelativeColumn().Stack(stack =>
                                {
                                    stack.Item().Padding(2).LineHorizontal(5).LineColor(Colors.Black);
                                    stack.Item().Text("Emergency Contact 1").FontColor("#397e89").FontSize(12).Bold();
                                    stack.Item().Padding(3).Text(employee.ec1_name).FontSize(9);
                                    stack.Item().Padding(3).Text(employee.ec1_relation).FontSize(9);
                                    stack.Item().Padding(3).Text(employee.ec1_address).FontSize(9);
                                    stack.Item().Text(employee.ec1_phone).FontSize(9);
                                });

                                // Emergency Contact 2 Column
                                row.RelativeColumn().Stack(stack =>
                                {
                                    stack.Item().Padding(2).LineHorizontal(5).LineColor(Colors.Black);
                                    stack.Item().Text("Emergency Contact 2").FontColor("#397e89").FontSize(12).Bold();
                                    stack.Item().Padding(3).Text(employee.ec2_name).FontSize(9);
                                    stack.Item().Padding(3).Text(employee.ec2_relation).FontSize(9);
                                    stack.Item().Padding(3).Text(employee.ec2_address).FontSize(9);
                                    stack.Item().Text(employee.ec2_phone).FontSize(9);
                                });
                            });




                            // Add spacing after each employee's section if needed
                            column.Item().PaddingBottom(5);
                        }


                    });
                    page.Footer().AlignCenter().Text("");
                });
            });

            // Function to get the path of the employee's image based on their ID
            string GetImagePath(string employeeId)
            {
                //Path to the images folder
                return $"C:/Users/rexja/source/repos/SoftwareTaskPDFGenerator/SoftwareTaskPDFGenerator/bin/debug/resources/{employeeId}.png";
            }


            string PDFfilePath = "C:/Users/rexja/OneDrive/Desktop/ST/Employees.pdf";
            document.GeneratePdf(PDFfilePath);

            Console.WriteLine($"PDF generated at {PDFfilePath}");
            Console.ReadKey();




        }
        
        // Function to read the CSV file and convert it to a list of employees
        static List<Employees> ReadEmployeesFromCsv(string filePath)
        {
            try
            {
                using (var reader = new StreamReader(filePath))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var records = csv.GetRecords<Employees>();
                    return records.ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading CSV file: {ex.Message}");
                return new List<Employees>();
            }
        }
    }
}
