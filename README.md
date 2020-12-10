# Angular10SpringBootExcelDownload

## Angular8/9/10 Code for Excel File Download feature from Front-End
#  Component.ts file add this code
downloadFile($event, element) {
    console.log("------DOWNLOAD CLICKED-----", $event, " :element= ", element);
    $event.stopPropagation();
    $event.preventDefault();
    console.log("------DOWNLOAD CLICKED logic start-----", $event);
    var fileName = "EmedNY_FullReport_" + element.transit_id + "-" + (new Date()).toISOString().replace(/-/g, '').split('T')[0] + ".xlsx";
    let postdata = element.transaction_id;
    this.commonService.showHideLoader("block");
    this.service.downloadExcelReport(postdata).subscribe(
      response => {
        this.commonService.showHideLoader("none");
        // var blob = new Blob([response._body], { type: 'application/vnd.ms-excel' });
        // var blob = new Blob([response._body], { type: 'application/octet-stream' });
        var contentType =  'application/vnd.ms-excel'; //'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        var blob = new Blob([response], { type: contentType });
        if (window.navigator && window.navigator.msSaveOrOpenBlob) {
          window.navigator.msSaveOrOpenBlob(blob, "fileName_test");
        } else {
          var a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          a.download = "fileName_test"+ ".xlsx";
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
        }
    }, error => {
      this.commonService.showHideLoader("none");
      console.log(error);
    });
  }
  
 # Service add this code
 @Injectable()
export class FileUploadService {

    constructor(private api: ApiService) {}

    upload(postdata: any): Observable<any> {
        return this.api.fileUpload(postdata);
    }

    getTransactionRecords(postdata: any): Observable<any> {
        return this.api.getTransactionRecord(postdata);
    }

    downloadExcelReport(postdata: any): Observable<any> {
        return this.api.downloadExcelReport(postdata);
    }
 }
 
 # Service code
     downloadExcelReport(data: any): Observable<any> {
        const MIME_TYPES = {
            xls: 'application/vnd.ms-excel',
            xlsx: 'application/vnc.openxmlformats-officedocument.spreadsheetml.sheet'
        }
        const httpOption: Object = {
            observe: 'response',
            // headers: new HttpHeaders({
            //   'Content-Type': 'application/json'
            // }),
            responseType: 'arraybuffer'
          };
        let headers = new HttpHeaders({
            'Content-Type': 'application/json'
         });
         let options = {
            headers: headers
         }

         let headers2 = new Headers();
            headers.append('Content-Type', 'application/json');
            headers.append('responseType', 'arrayBuffer');
        return this.httpClient.post(this.BASE_URL + "/api/v1/file/download", data, { responseType: 'blob'})
                    // .map((response) => {
                    //     return new Blob([response.toString()], { type: 'application/vnd.ms-excel' });
                    // })
                    .catch(this.handleError);
    }
    
    
   # Spring Boot Code Snippet
  SomeController.java
  
  @PostMapping
	@RequestMapping("/download")
	public ResponseEntity<?> exportExcelReport(HttpServletRequest request, HttpServletResponse response, @RequestBody String transit_id) {
		XSSFWorkbook workbook = null;
		byte[] contentReturn = null;
		ByteArrayInputStream in = null;
		String fileName = "";
		//		HttpHeaders headers = null;
		try {
			/**
			 * Get 271 Data From DB
			 */

			/* Logic to Export Excel */
			LocalDateTime localDate = LocalDateTime.now();
			//			EmedNY_FullReport_5991_20200728193019PM
			fileName = "EmedNY_FullReport_" + transit_id + "-" + localDate.toString() + ".xlsx";

			// Pass data received from DB to below export function
			workbook = (XSSFWorkbook) fileUploadService.exportToExcel();
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			workbook.write(baos);
			contentReturn = baos.toByteArray();
			in = new ByteArrayInputStream(baos.toByteArray());

			/* Export Excel logic end */

		} catch (Exception ecx) {
			return ResponseEntity.status(HttpStatus.SC_INTERNAL_SERVER_ERROR).body(ecx.getCause());
		} finally {
			if (null != workbook) {
				try {
					workbook.close();
				} catch (IOException eio) {
					System.out.println("Error Occurred while exporting to XLS ");
					eio.printStackTrace();
				}
			}
		}
		HttpHeaders headers = new HttpHeaders();
		headers.add("Content-Disposition", "inline: fileName=fileName_test.xlsx");
		return ResponseEntity.ok().headers(headers).body(new InputStreamResource(in));

	}
  
  # Check fileUploadService.exportToExcel(); class and other Util class for this functionality
  # FileUploadService.java
  package com.evt.dsnp.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import com.evt.dsnp.pojo.Customer;
import com.evt.dsnp.repository.MemberFileUploadRepository;
import com.evt.dsnp.util.CVSUtils;

@Component
public class FileUploadService {
	@Autowired
	private MemberFileUploadRepository fileUploadRepo;

	public MemberFileUploadRepository getFileUploadRepo() {
		return fileUploadRepo;
	}

	public Object exportToExcel() {
		List<String> data = null;
//		https://anupampawar.com/2019/03/08/angular-and-spring-boot-excel-export-download/
//		https://grokonez.com/spring-framework/spring-boot/excel-file-download-from-springboot-restapi-apache-poi-mysql
		String[] columns = setUpExcelColumns();
//		Workbook workbook = CVSUtils.writeExcelFile(columns, summaryArray);
		Workbook workbook = CVSUtils.writeExcelFile(columns, data);
		return workbook;
	}

	private String[] setUpExcelColumns() {
		String[] cols = {"CIN#", "DOB", "GENDER", "COUNTY", "DATE OF SERVICE", "USERID#", "USER INFO",
				"INSURANCE LEVEL CODE", "SERVICE TYPE CODE"};
		return cols;
	}

	public static ByteArrayInputStream contactListToExcelFile(List<Customer> customers) {
		try(Workbook workbook = new XSSFWorkbook()){
			Sheet sheet = workbook.createSheet("Customers");
			
			Row row = sheet.createRow(0);
	        CellStyle headerCellStyle = workbook.createCellStyle();
	        headerCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
	        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        // Creating header
	        Cell cell = row.createCell(0);
	        cell.setCellValue("First Name");
	        cell.setCellStyle(headerCellStyle);
	        
	        cell = row.createCell(1);
	        cell.setCellValue("Last Name");
	        cell.setCellStyle(headerCellStyle);
	
	        cell = row.createCell(2);
	        cell.setCellValue("Mobile");
	        cell.setCellStyle(headerCellStyle);
	
	        cell = row.createCell(3);
	        cell.setCellValue("Email");
	        cell.setCellStyle(headerCellStyle);
	        
	        // Creating data rows for each customer
	        for(int i = 0; i < customers.size(); i++) {
	        	Row dataRow = sheet.createRow(i + 1);
	        	dataRow.createCell(0).setCellValue(customers.get(i).getFirstName());
	        	dataRow.createCell(1).setCellValue(customers.get(i).getLastName());
	        	dataRow.createCell(2).setCellValue(customers.get(i).getMobileNumber());
	        	dataRow.createCell(3).setCellValue(customers.get(i).getEmail());
	        }
	
	        // Making size of column auto resize to fit with data
	        sheet.autoSizeColumn(0);
	        sheet.autoSizeColumn(1);
	        sheet.autoSizeColumn(2);
	        sheet.autoSizeColumn(3);
	        
	        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
	        workbook.write(outputStream);
	        return new ByteArrayInputStream(outputStream.toByteArray());
		} catch (IOException ex) {
			ex.printStackTrace();
			return null;
		}
	}

}

# CVSUtils.java
package com.evt.dsnp.util;

import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CVSUtils {


	public static Workbook writeExcelFile(String[] columns, List<String> data) {

		Workbook workbook = new XSSFWorkbook();
		int rowNum = 0;
		// Create a Font for styling header cells
		Font headerFont = setHeaderFont(workbook);

		// Create a CellStyle with the font
		CellStyle headerCellStyle = setHeaderCellStyle(workbook, headerFont);

		// Create a CellStyle for CampaignSummary main object with Grey Color
		CellStyle styleGrey = workbook.createCellStyle();
		styleGrey.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
//		styleGrey.setFillPattern(CellStyle.SOLID_FOREGROUND);

		Sheet reportSheet = workbook.createSheet("FullMemberReport");

		//Row for header
		Row headerRow = reportSheet.createRow(0);

		//Header
		createHeaders(columns, headerCellStyle, headerRow);

		// CellStyle for Age
		//	      CellStyle ageCellStyle = workbook.createCellStyle();
		//	      ageCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("#"));

		//Create the 'report summary sheet': setCellValue
		createReportSummary(rowNum, reportSheet, columns);

		return workbook;

	}

	private static void createHeaders(String[] columns, CellStyle headerCellStyle, Row headerRow) {
		for (int col = 0; col < columns.length; col++) {
			Cell cell = headerRow.createCell(col);
			cell.setCellValue(columns[col]);
			cell.setCellStyle(headerCellStyle);
		}
	}

	private static CellStyle setHeaderCellStyle(Workbook workbook, Font headerFont) {
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);
		return headerCellStyle;
	}

	private static Font setHeaderFont(Workbook workbook) {
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setColor(IndexedColors.BLUE.getIndex());
		return headerFont;
	}

	private static void createReportSummary(int rowNum, Sheet reportSummarySheet, String[] columns) {
		Row vehDetailsRow = reportSummarySheet.createRow(++rowNum);
		//		CVSSummary cvsSummary = vehicleLastSeenByCampaignReport.getCvsSummary();
		int columnNum = 0;

		vehDetailsRow.createCell(columnNum).setCellValue(columnNum);
		vehDetailsRow.createCell(++columnNum).setCellValue(++columnNum+"-Test");
		vehDetailsRow.createCell(++columnNum).setCellValue(++columnNum+"-Test");
		vehDetailsRow.createCell(++columnNum).setCellValue(++columnNum+"-Test");     
		vehDetailsRow.createCell(++columnNum).setCellValue(++columnNum+"-Test");

		for (int i = 0; i < columns.length; i++) {
			reportSummarySheet.autoSizeColumn(i);
		}
	}
}
