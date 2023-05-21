package ru.kaleev.SpringCourse.controllers;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;
import org.springframework.data.domain.Sort;
import org.springframework.data.domain.Sort.Direction;
import org.springframework.data.domain.Sort.Order;
import org.springframework.data.domain.PageImpl;
import org.springframework.security.access.prepost.PreAuthorize;
import org.springframework.security.authentication.AuthenticationManager;
import org.springframework.security.authentication.UsernamePasswordAuthenticationToken;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContext;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import ru.kaleev.SpringCourse.entity.User;
import ru.kaleev.SpringCourse.repository.UserRepository;
import ru.kaleev.SpringCourse.service.impl.ExportSalesExcel;
import ru.kaleev.SpringCourse.service.impl.PdfPageServiceImpl;

import org.springframework.ui.ModelMap;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

@Controller
@RequestMapping("/login")
public class FirstController {
	@Autowired
	ExportSalesExcel exportSalesExcel;
	
	@Autowired
	AuthenticationManager authenticationManager;

	@Autowired
	UserRepository userRepository;
	
	@Autowired
	PdfPageServiceImpl pdfPageServiceImpl;

	@GetMapping("/signin")
	public String signin(HttpSession session,
			ModelMap model) throws IOException {
		
		Object errorMessage = session.getAttribute("errorMessage");
		if (errorMessage != null) {
			model.addAttribute("errorMessage", errorMessage);
			session.removeAttribute("errorMessage");
		}

		Object checkAuthor = session.getAttribute("checkAuthor");
		if (checkAuthor != null) {
			model.addAttribute("checkAuthor", checkAuthor);
			session.removeAttribute("checkAuthor");
		}
		
	    session.invalidate();

		return "first/loginPage";

	}

	@PostMapping("/success")
	public String authenticate(@RequestParam(value = "username", required = false) String username,
			@RequestParam(value = "password", required = false) String password,
			HttpSession session) {

		Authentication authentication = authenticationManager
				.authenticate(new UsernamePasswordAuthenticationToken(username, password));
		
		if (authentication.isAuthenticated()) {
			SecurityContext securityContext = SecurityContextHolder.getContext();
			securityContext.setAuthentication(authentication);

			return "redirect:/login/success";
		} else {
			session.setAttribute("errorMessage", "Invalid username or password");
			return "redirect:/login/signin";
		}
	}

    @PreAuthorize("hasAuthority('ADMIN')")
	@GetMapping("/success")
	public String successPage(ModelMap model, @RequestParam(required = false) String keyword,
											@RequestParam(defaultValue = "1") int page,
											@RequestParam(defaultValue = "4") int size,
											@RequestParam(defaultValue = "id,asc") String[] sort) {
    	List<User> listUser = new ArrayList<User>();
    	
    	String sortField = sort[0];
        String sortDirection = sort[1];    
        Direction direction = sortDirection.equals("desc") ? Sort.Direction.DESC : Sort.Direction.ASC;
        Order order = new Order(direction, sortField);
    	
        Pageable pageable = PageRequest.of(page-1, size, Sort.by(order));
        Page<User> userPage;
        if (keyword == null) {
        	userPage = userRepository.findAll(pageable);
        
          } else {
        	  userPage = userRepository.findByAuthorContaining(keyword, pageable);
        	  model.addAttribute("keyword", keyword);
          }
        
        listUser = userPage.getContent();
        
        model.addAttribute("tutorials", listUser);
        model.addAttribute("currentPage", userPage.getNumber() + 1); 	
        model.addAttribute("totalPages", userPage.getTotalPages()); 	
        model.addAttribute("totalItems", userPage.getTotalElements()); 	
        model.addAttribute("pageSize", size);
        model.addAttribute("sortField", sortField);
        model.addAttribute("sortDirection", sortDirection);
        model.addAttribute("reverseSortDirection", sortDirection.equals("asc") ? "desc" : "asc");
    	        
		return "first/hello";
	}
    
    @GetMapping(value = "/exportExcel")
    public void exportExcel(HttpServletResponse response, @RequestParam(required = false) String keyword,
    		@RequestParam(defaultValue = "1") int page,
    		@RequestParam(defaultValue = "4") int size,
    		@RequestParam(defaultValue = "id,asc") String[] sort) throws Exception {
//    	List<User> listUser = new ArrayList<User>();
//    	
//    	String sortField = sort[0];
//    	String sortDirection = sort[1];
//    	Direction direction = sortDirection.equals("desc") ? Sort.Direction.DESC : Sort.Direction.ASC;
//    	Order order = new Order(direction, sortField);
//    	
//    	Pageable pageable = PageRequest.of(page-1, size, Sort.by(order));
//    	Page<User> userPage;
//    	if (keyword == null) {
//    		userPage = userRepository.findAll(pageable);
//    	} else {
//    		userPage = userRepository.findByAuthorContaining(keyword, pageable);
//    	}
//    	
//    	listUser = userPage.getContent();
//    	
//    	// Create Excel workbook and sheet using Apache POI
//    	XSSFWorkbook workbook = new XSSFWorkbook();
//    	XSSFSheet sheet = workbook.createSheet("User List");
//    	
//    	// Create header row
//    	XSSFRow headerRow = sheet.createRow(0);
//    	XSSFCellStyle headerCellStyle = workbook.createCellStyle();
//    	XSSFFont headerFont = workbook.createFont();
//    	headerFont.setBold(true);
//    	headerCellStyle.setFont(headerFont);
//    	
//    	XSSFCell idHeaderCell = headerRow.createCell(0);
//    	idHeaderCell.setCellValue("ID");
//    	idHeaderCell.setCellStyle(headerCellStyle);
//    	
//    	XSSFCell nameHeaderCell = headerRow.createCell(1);
//    	nameHeaderCell.setCellValue("Name");
//    	nameHeaderCell.setCellStyle(headerCellStyle);
//    	
//    	XSSFCell roleHeaderCell = headerRow.createCell(2);
//    	roleHeaderCell.setCellValue("Author");
//    	roleHeaderCell.setCellStyle(headerCellStyle);
//    	
//    	// Create data rows
//    	XSSFCellStyle cellStyle = workbook.createCellStyle();
//    	XSSFFont cellFont = workbook.createFont();
//    	cellFont.setBold(false);
//    	cellStyle.setFont(cellFont);
//    	
//    	int rowNum = 1;
//    	for (User user : listUser) {
//    		XSSFRow row = sheet.createRow(rowNum++);
//    		
//    		XSSFCell idCell = row.createCell(0);
//    		idCell.setCellValue(user.getId());
//    		idCell.setCellStyle(cellStyle);
//    		
//    		XSSFCell nameCell = row.createCell(1);
//    		nameCell.setCellValue(user.getUsername());
//    		nameCell.setCellStyle(cellStyle);
//    		
//    		XSSFCell roleCell = row.createCell(2);
//    		roleCell.setCellValue(user.getAuthor());
//    		roleCell.setCellStyle(cellStyle);
//    	}
//    	
//    	// Autosize columns
//    	sheet.autoSizeColumn(0);
//    	sheet.autoSizeColumn(1);
//    	sheet.autoSizeColumn(2);
    	
    	// Set the content type and attachment header
    	response.setContentType("application/vnd.ms-excel");
    	response.setHeader("Content-Disposition", "attachment; filename=user_list.xlsx");
    	
    	// Write the Excel to the response output stream
    	OutputStream outputStream = response.getOutputStream();
//    	workbook.write(outputStream);
    	ByteArrayOutputStream byteArrayOutputStream = exportSalesExcel.exportExcel(null);//
    	byteArrayOutputStream.writeTo(outputStream);//
    	outputStream.flush();
    	outputStream.close();
    }
	
    @GetMapping(value = "/exportPDF")
    public void exportPDF(HttpServletResponse response, @RequestParam(required = false) String keyword,
                          @RequestParam(defaultValue = "1") int page,
                          @RequestParam(defaultValue = "4") int size,
                          @RequestParam(defaultValue = "id,asc") String[] sort) throws Exception {
        List<User> listUser = new ArrayList<User>();

        String sortField = sort[0];
        String sortDirection = sort[1];
        Direction direction = sortDirection.equals("desc") ? Sort.Direction.DESC : Sort.Direction.ASC;
        Order order = new Order(direction, sortField);

        Pageable pageable = PageRequest.of(page-1, size, Sort.by(order));
        Page<User> userPage;
        if (keyword == null) {
            userPage = userRepository.findAll(pageable);
        } else {
            userPage = userRepository.findByAuthorContaining(keyword, pageable);
        }

        listUser = userPage.getContent();

        // Create PDF document using iText
        Document document = new Document(PageSize.A4);
        ByteArrayOutputStream baosPDF = new ByteArrayOutputStream();
        PdfWriter writer = PdfWriter.getInstance(document, baosPDF);
        document.open();
        document.add(new Paragraph("User List"));

        PdfPTable table = new PdfPTable(3);
        table.setWidthPercentage(100);
        table.setSpacingBefore(10f);
        table.setSpacingAfter(10f);

        Font headerFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD);
        PdfPCell headerCell1 = new PdfPCell(new Phrase("ID", headerFont));
        PdfPCell headerCell2 = new PdfPCell(new Phrase("Name", headerFont));
        PdfPCell headerCell3 = new PdfPCell(new Phrase("Author", headerFont));
        table.addCell(headerCell1);
        table.addCell(headerCell2);
        table.addCell(headerCell3);

        Font cellFont = FontFactory.getFont(FontFactory.HELVETICA);
        for (User user : listUser) {
            table.addCell(new PdfPCell(new Phrase(String.valueOf(user.getId()), cellFont)));
            table.addCell(new PdfPCell(new Phrase(user.getUsername(), cellFont)));
            table.addCell(new PdfPCell(new Phrase(user.getAuthor(), cellFont)));
        }

        document.add(table);
        document.close();

        // Set the content type and attachment header
        response.setContentType("application/pdf");
        response.setHeader("Content-Disposition", "attachment; filename=user_list.pdf");

        // Write the PDF to the response output stream
        OutputStream outputStream = response.getOutputStream();
        baosPDF.writeTo(outputStream);
        outputStream.flush();
        outputStream.close();
    }
    
    @GetMapping(value = "/exportPDFAll")
    public void exportPDFAll(HttpServletResponse response, @RequestParam(required = false) String keyword,
    		@RequestParam(defaultValue = "1") int page,
    		@RequestParam(defaultValue = "4") int size,
    		@RequestParam(defaultValue = "id,asc") String[] sort) throws Exception {
    	
    	pdfPageServiceImpl.exportPDFAll(response, keyword, page, size, sort);
    }
    
    @GetMapping(value = "/exportPDFByHTML")
    public void exportPDFByHTML(HttpServletRequest request, HttpServletResponse response,
    		@RequestParam(required = false) String keyword,
    		@RequestParam(defaultValue = "1") int page,
    		@RequestParam(defaultValue = "4") int size,
    		@RequestParam(defaultValue = "id,asc") String[] sort) throws Exception {
    	
    	pdfPageServiceImpl.generatePdfFromHtml(request, response, "first/renderPDF", keyword, 1, size, sort);
    }
    
	@GetMapping("/logout")
	public String logout(HttpServletRequest request) {
		return null;
		/*
		 * request.getSession().invalidate(); return "redirect:/login/signin";
		 */
	}
	
}
