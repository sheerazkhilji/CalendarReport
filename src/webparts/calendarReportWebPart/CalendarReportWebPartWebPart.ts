import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { jsPDF } from 'jspdf'; // Import jsPDF directly
import 'jspdf-autotable'; // Import the autotable plugin
import * as strings from 'CalendarReportWebPartWebPartStrings';
import 'bootstrap/dist/css/bootstrap.min.css';

import styles from './CalendarReportWebPartWebPart.module.scss';
import axios from 'axios';

export interface ICalendarReportWebPartWebPartProps {
  description: string;
}

export default class CalendarReportWebPartWebPart extends BaseClientSideWebPart<ICalendarReportWebPartWebPartProps> {



  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.calendarReport}">
        <h2 class="mb-4">Training Calendar</h2>
        <div class="row">
          <div class="col-md-6">
            <div class="mb-3">
              <label for="department" class="form-label">Select Department</label>
              <select class="form-select" id="department">
              </select>
            </div>
            <div class="mb-3">
              <label for="month" class="form-label">Month</label>
              <select class="form-select" id="month">
               <option value="January">January</option>
                   <option value="February">February</option>
                   <option value="March">March</option>
                   <option value="April">April</option>
                   <option value="May">May</option>
                   <option value="June">June</option>
                   <option value="July">July</option>
                   <option value="August">August</option>
                   <option value="September">September</option>
                   <option value="October">October</option>
                   <option value="November">November</option>
                   <option value="December">December</option>
              </select>
            </div>
            <div class="mb-3">
              <label for="typeOfTraining" class="form-label">Type of Training</label>
              <select class="form-select" id="typeOfTraining">
                <option value="SOP">SOP</option>
                  <option value="GMP">GMP</option>
                    <option value="GMP & GSP">GMP & GSP</option>
                      <option value="Awareness Session">Awareness Session</option>
              </select>
            </div>
            <div class="mb-3">
              <label for="preparedBy" class="form-label">Prepared By</label>
              <input type="text" class="form-control" id="preparedBy" value="SYED IBRAHIM">
            </div>
            <div class="mb-3">
              <label for="reviewedBy" class="form-label">Reviewed By</label>
              <input type="text" class="form-control" id="reviewedBy" value="AZHAR ZUBERI">
            </div>
          </div>

          <div class="col-md-6">
            <div class="mb-3">
              <label for="year" class="form-label">Year</label>
          
              <select class="form-select" id="year">
              <option value="2024">2024</option>
                   <option value="2025">2025</option>
                   <option value="2026">2026</option>
                   <option value="2027">2027</option>
                   <option value="2028">2028</option>
                   <option value="2029">2029</option>
                   <option value="2030">2030</option>
              
              </select>
            </div>
            <div class="mb-3">
              <label for="natureOfTraining" class="form-label">Nature of Training</label>
              <select class="form-select" id="natureOfTraining">
                <option value="Additional Training">Additional Training</option>

                 <option value="Plan Training">Plan Training</option>
              </select>
            </div>
            <div class="mb-3">
              <label for="reportFooter" class="form-label">Report With Footer</label>
              <select class="form-select" id="reportFooter">
                <option value="Yes">Yes</option>
                 <option value="No">No</option>
              </select>
            </div>
            <div class="mb-3">
              <label for="checkedBy" class="form-label">Checked By</label>
              <input type="text" class="form-control" id="checkedBy" value="SOHAIL TARIQ">
            </div>
            <div class="mb-3">
              <label for="approvedBy" class="form-label">Approved By</label>
              <input type="text" class="form-control" id="approvedBy" value="SYED IBRAHIM">
            </div>
          </div>
        </div>

        <button id="fetchButton" class="btn btn-primary">Fetch</button>
        <button id="pdfButton" class="btn btn-danger">Download PDF</button>
      </div>

      <div class="header">
        <h2>GETZ PHARMA (PRIVATE) LIMITED</h2>
        <p>Annual Training Calendar</p>
        <p>Department: <span id="reportDepartment"></span> | Year: <span id="reportYear"></span></p>

      </div>

      <table id="reportTable" style="display: none;" class="${styles.table}">
        <thead>
          <tr>
            <th>Nature of Training</th>
            <th>Type of Training</th>
            <th>Subject</th>
            <th>Reference</th>
            <th>Planned Month</th>
            <th>Trainer's Name</th>
            <th>Training Date</th>
            <th>Duration (Minutes)</th>
            <th>Total Trainee</th>
            <th>Total Trainer</th>
            <th>Total Attendees</th>
            <th>Total Hours</th>
            <th>Remarks</th>
          </tr>
        </thead>
        <tbody>
          <!-- Rows will be dynamically inserted here -->
        </tbody>
      </table>
    `;
    this.populateDepartmentDropdown();
    this.addEventListeners();
  }



// Method to populate the Department dropdown
private async populateDepartmentDropdown(): Promise<void> {
  try {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const departmentListUrl = `${siteUrl}/_api/web/lists/getByTitle('Department')/items?$select=Title`;
    debugger;
    // Fetch the department list items using Axios
    const response = await axios.get(departmentListUrl, {
      headers: {
        "Accept": "application/json;odata=verbose",
        "odata-version": "",
        "IF-MATCH": "*",
     
      }
    });

    const departments = response.data.d.results;
    const departmentSelect = this.domElement.querySelector("#department") as HTMLSelectElement;
debugger;
    // Clear existing options
    departmentSelect.innerHTML = '<option value="">Select Department</option>';

    // Populate the dropdown with items from the Department list
    departments.forEach((dept: { Title: string; }) => {
      const option = document.createElement("option");
      option.value = dept.Title;
      option.text = dept.Title;
      departmentSelect.appendChild(option);
    });
  } catch (error) {
    console.error("Error fetching departments:", error);
  }
}






  private addEventListeners(): void {
    document.getElementById('fetchButton')?.addEventListener('click', () => this.fetchData());
    document.getElementById('pdfButton')?.addEventListener('click', () => this.downloadPDF());
  }

  private  async fetchData(): Promise<void> {

    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      
      const Department = (this.domElement.querySelector("#department") as HTMLSelectElement).value;
      const Month = (this.domElement.querySelector("#month") as HTMLSelectElement).value;
      const TypeofTraining = (this.domElement.querySelector("#typeOfTraining") as HTMLSelectElement).value;
      const Year = (this.domElement.querySelector("#year") as HTMLSelectElement).value;
      const NatureofTraining = (this.domElement.querySelector("#natureOfTraining") as HTMLSelectElement).value;
  
      // Construct API URL with filter conditions
      let filterConditions = [];
      if (Department) filterConditions.push(`Department eq '${Department}'`);
      if (Month) filterConditions.push(`Month eq '${Month}'`);
      if (TypeofTraining) filterConditions.push(`TypeofTraining eq '${TypeofTraining}'`);
      if (Year) filterConditions.push(`Year eq '${Year}'`);
      if (NatureofTraining) filterConditions.push(`NatureofTraining eq '${NatureofTraining}'`);
      
      const filterQuery = filterConditions.length ? `$filter=${filterConditions.join(" and ")}` : "";
      const apiUrl = `${siteUrl}/_api/web/lists/getByTitle('Annual Training Calendar')/items?${filterQuery}`;
  
      const response = await axios.get(apiUrl, {
        headers: {
           "Accept": "application/json;odata=verbose",
          "odata-version": "",
          "IF-MATCH": "*",
        }
      });
  
      const Resdata = response.data.d.results;
     

      // Get form values
      const departmentElement = document.getElementById('department') as HTMLSelectElement | null;
      const yearElement = document.getElementById('year') as HTMLInputElement | null;
  
      if (departmentElement && yearElement) {
        const department = departmentElement.value;
        const year = yearElement.value;
  
        // Populate the header with dynamic department and year
        const reportDepartmentElement = document.getElementById('reportDepartment');
        const reportYearElement = document.getElementById('reportYear');
  
        if (reportDepartmentElement && reportYearElement) {
          reportDepartmentElement.innerText = department;
          reportYearElement.innerText = year;
        }
  
        // Clear existing rows
        const tableBody = document.querySelector('#reportTable tbody') as HTMLTableSectionElement | null;
        if (tableBody) {
          tableBody.innerHTML = '';
  
          // Populate the table with data
          Resdata.forEach((row: { NatureofTraining: any; TypeofTraining: any; Subject : any; Reference: any; Month: any; TrainerName : any; Department: any; Duration_x0028_InMinutes_x0029_: any; TotalTrainer : any; TotalAttendees: any; TotalTrainee : any; TotalHours: any; remarks : any; }) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
           <td>${row?.NatureofTraining || ''}</td>
<td>${row?.TypeofTraining || ''}</td>
<td>${row?.Subject || ''}</td>
<td>${row?.Reference || ''}</td>
<td>${row?.Month || ''}</td>
<td>${row?.TrainerName || ''}</td>
<td>${row?.Department || ''}</td>
<td>${row?.Duration_x0028_InMinutes_x0029_ || ''}</td>
<td>${row?.TotalTrainer || ''}</td>
<td>${row?.TotalAttendees || ''}</td>
<td>${row?.TotalTrainee || ''}</td>
<td>${row?.TotalHours || ''}</td>
<td>${row?.remarks || ''}</td>

            `;
            tableBody.appendChild(tr);
          });
  
          // Show the table and header once data is fetched
          const reportTableElement = document.getElementById('reportTable') as HTMLTableElement | null;
          const headerElement = document.querySelector('.header') as HTMLElement | null;
  
          if (reportTableElement && headerElement) {
            reportTableElement.style.display = 'table';
            headerElement.style.display = 'block';
          }
        }
      }



    } catch (error) {
      console.error("Error fetching filtered data:", error);
    }


 
  }


  

  // Function to fetch the image URL from SharePoint
  private async fetchImageFromSharePoint(url: string): Promise<string | null> {
    try {
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          Accept: 'application/json;odata=verbose',
        },
      });

      if (!response.ok) {
        throw new Error('Network response was not ok');
      }

      const data = await response.json();
      const imageFileRef = data.d.results[0]?.FileRef; // Get the FileRef of the first image

      if (imageFileRef) {
        // Construct the full URL
        const fileUrl = `${this.context.pageContext.web.absoluteUrl}${imageFileRef}`;
        
        return fileUrl;
      }
    } catch (error) {
      console.error('Error fetching image:', error);
    }
    return null; // Return null if there was an error
  }

  private async downloadPDF(): Promise<void> {
    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a3',
      putOnlyUsedFonts: true,
      floatPrecision: 16,
    });
    const libraryName = 'GetzLogo'; // Update to your document library name if needed
    const query = `/_api/web/lists/getbytitle('${libraryName}')/items?$select=FileRef,FileLeafRef&$filter=substringof('.jpeg',FileLeafRef) or substringof('.png',FileLeafRef) or substringof('.gif',FileLeafRef)`;

    // Fetch the logo image URL from SharePoint
    const logoUrl = await this.fetchImageFromSharePoint(this.context.pageContext.web.absoluteUrl + query);

    // Function to add the header to the document
    const addHeader =  (doc: any) => {
 
      // Check if the image was successfully fetched
      if (logoUrl) {
        const logo = logoUrl.replace(/\/TestSite(\/)?/, '/');
  
        // Set the fill color and draw header
        doc.setFillColor(135, 206, 235); // Sky blue color
        doc.rect(15, 10, 390, 27, 'FD'); // Header rectangle
  
        // Company name and logo
        doc.setFontSize(16);
        doc.setFont("helvetica", "bold");
         doc.addImage(logo, 'JPEG', 20, 15, 15, 15); // Logo
        doc.text("GETZ PHARMA (PRIVATE) LIMITED", 160, 20);
  
        // Document info
        doc.rect(345, 12, 50, 24); // Document info rectangle
        doc.setFontSize(8);
        doc.text("Form #: F-QA-050", 350, 15);
        doc.text("Rev #: 10", 350, 20);
        doc.text("SOP Ref #: QA-030", 350, 25);
        doc.text("Change alert #: 9007206", 350, 30);
        doc.text("Issued Date: 02/08/2023", 350, 35);
  
        // Department and Year
        doc.setFontSize(12);
        const department = document.getElementById('reportDepartment')?.innerText || 'Please Select';
        const year = document.getElementById('reportYear')?.innerText || 'Please Select';
        doc.setFont("helvetica", "bold");
        doc.text("Annual Training Calendar", 180, 30); // Title
  
        // Department/year rectangle
        doc.setFillColor(135, 206, 235); // Sky blue color
        doc.rect(15, 40, 390, 20, 'FD');
        doc.text(`Department: ${department}`, 22, 45);
        doc.text(`Year: ${year}`, 185, 45);
      } else {
        console.warn('Logo URL could not be fetched.');
      }
    };
    const drawBlocks1 = (doc: any) => {
      const startX = 15; // Start blocks from 20mm on X-axis
      const pageWidth = 405; // A3 landscape width in mm
      const blockHeight = 8; // Height of each block
      const blockWidth = (pageWidth - startX) / 4; // Adjust the width to fit 4 blocks from X = 20mm
  
      // Loop to draw blocks
      for (let i = 0; i < 4; i++) {
        const xPosition = startX + i * blockWidth; // Calculate X position for each block
        const yPosition = 260; // Y position (can be adjusted)
  
        // Draw the rectangle
        doc.rect(xPosition, yPosition, blockWidth, blockHeight, 'S'); // 'S' draws the border
  
      }
    };
    // Function to draw footer blocks
    const drawBlocks = (doc: any) => {
      const blockTexts = ['Prepared by', 'Checked by', 'Approved by', 'Reviewed by'];
      const startX = 15; 
      const pageWidth = 405; 
      const blockHeight = 5; 
      const blockWidth = (pageWidth - startX) / 4; 
  
      for (let i = 0; i < 4; i++) {
        const xPosition = startX + i * blockWidth; 
        const yPosition = 270; 
        doc.rect(xPosition, yPosition, blockWidth, blockHeight, 'S'); // Draw the border
        doc.setFontSize(10);
        const text = blockTexts[i]; 
        const textWidth = doc.getTextWidth(text);
        const textXPosition = xPosition + (blockWidth - textWidth) / 2; 
        doc.text(text, textXPosition, yPosition + (blockHeight / 1.5));
      }
    };
  
    // Function to add footer on each page
    const addFooter = (doc: any, pageNumber: number, pageCount: number) => {
      drawBlocks1(doc);
      drawBlocks(doc);
      doc.rect(15, 278, 390, 10); 
      doc.setTextColor(255, 0, 0);
      doc.setFont("helvetica", "italic");
      doc.text('This Plan is tentative, in case of any change in the training plan ', 20, 283); 
      doc.text('be mentioned in remarks ', 50, 286.5); 
      doc.setTextColor(0, 0, 0);
      doc.setFontSize(10);
      doc.text(`Page ${pageNumber} of ${pageCount}`, 380, 287); 
    };
  


    const siteUrl = this.context.pageContext.web.absoluteUrl;
      
    const Department = (this.domElement.querySelector("#department") as HTMLSelectElement).value;
    const Month = (this.domElement.querySelector("#month") as HTMLSelectElement).value;
    const TypeofTraining = (this.domElement.querySelector("#typeOfTraining") as HTMLSelectElement).value;
    const Year = (this.domElement.querySelector("#year") as HTMLSelectElement).value;
    const NatureofTraining = (this.domElement.querySelector("#natureOfTraining") as HTMLSelectElement).value;
    const Isfooter = (this.domElement.querySelector("#reportFooter") as HTMLSelectElement).value;
    // Construct API URL with filter conditions
    let filterConditions = [];
    if (Department) filterConditions.push(`Department eq '${Department}'`);
    if (Month) filterConditions.push(`Month eq '${Month}'`);
    if (TypeofTraining) filterConditions.push(`TypeofTraining eq '${TypeofTraining}'`);
    if (Year) filterConditions.push(`Year eq '${Year}'`);
    if (NatureofTraining) filterConditions.push(`NatureofTraining eq '${NatureofTraining}'`);
    
    const filterQuery = filterConditions.length ? `$filter=${filterConditions.join(" and ")}` : "";
    const apiUrl = `${siteUrl}/_api/web/lists/getByTitle('Annual Training Calendar')/items?${filterQuery}`;

    const response = await axios.get(apiUrl, {
      headers: {
         "Accept": "application/json;odata=verbose",
        "odata-version": "",
        "IF-MATCH": "*",
      }
    });

    const Resdata = response.data.d.results;
   
var data: { nature: any; type: any; subject: any; reference: any; month: any; trainer: any; Department: any; duration: any; trainee: any; trainer_count: any; attendees: any; hours: any; remarks: any; }[]=[];

    Resdata.forEach((row: { NatureofTraining: any; TypeofTraining: any; Subject : any; Reference: any; Month: any; TrainerName : any; Department: any; Duration_x0028_InMinutes_x0029_: any; TotalTrainer : any; TotalAttendees: any; TotalTrainee : any; TotalHours: any; remarks : any; }) => {
    
      data.push(

        {
          nature: row?.NatureofTraining || '',
          type: row?.TypeofTraining || '',
          subject: row?.Subject || '',
          reference: row?.Reference || '',
          month: row?.Month || '',
          trainer: row?.TrainerName || '',
          Department: row?.Department || '',
          duration: row?.Duration_x0028_InMinutes_x0029_ || '',
          trainee: row?.TotalTrainee || '',
          trainer_count: row?.TotalTrainer || '',
          attendees: row?.TotalAttendees || '',
          hours: row?.TotalHours || '',
          remarks: row?.remarks || '',
          
        }
      )

    
    });



  
    // Track the current page number manually
  
    // Add header for the first page
     addHeader(doc);
  
     (doc as any).autoTable({
      head: [['Nature', 'Type', 'Subject', 'Reference', 'Month', 'Trainer', 'Department', 'Duration', 'Trainees', 'Trainers', 'Attendees', 'Hours', 'Remarks']],
      body: data.map(item => [
          item.nature, item.type, item.subject, item.reference, item.month,
          item.trainer, item.Department, item.duration, item.trainee, item.trainer_count,
          item.attendees, item.hours, item.remarks,
      ]),
      startY: 65,
      theme: 'grid',
      margin: { top: 65, bottom: 35 },
      headStyles: {
          fillColor: [135, 206, 235], // Sky blue color
          textColor: [0, 0, 0], // Black text
          fontStyle: 'bold', // Optional: Make the header text bold
      },
      didDrawPage: function (data: any) {
          // Add header to each page
          addHeader(doc);
  
          // Add footer for each page
          const pageCount = doc.internal.pages.length - 1; // Adjusts to exclude empty pages
          const pageNumber = 1; // Get the current page number
      
          if(Isfooter=="Yes"){

            addFooter(doc, pageNumber, pageCount);
          }
      },
  });
  
  
    // Save the PDF
    doc.save('Training_Calendar.pdf');
  }
  



  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
