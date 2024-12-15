package org.sample.BTCK;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.spire.pdf.PdfDocument;
import com.spire.pdf.FileFormat;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class App {
    public static void main(String[] args) throws IOException {
        // Đường dẫn đến file PDF gốc
        File pdffile = new File("D:\\2024\\Java\\slide\\Bai_Giang_Lap_Trinh_Java-Hamv1.pdf");
        PDDocument document = PDDocument.load(pdffile);

        int pagesPerSplit = 10; // Số trang mỗi phần PDF
        int totalPages = document.getNumberOfPages();

        // Danh sách các file DOCX để nối
        String[] docxFiles = new String[(totalPages / pagesPerSplit) + (totalPages % pagesPerSplit > 0 ? 1 : 0)];

        // Tách PDF thành các phần nhỏ, mỗi phần chứa 10 trang và chuyển thành DOCX
        for (int i = 0; i < totalPages; i += pagesPerSplit) {
            // Tạo một tài liệu PDF mới cho mỗi phần
            PDDocument splitDocument = new PDDocument();
            for (int j = i; j < i + pagesPerSplit && j < totalPages; j++) {
                PDPage page = document.getPage(j); // Lấy trang thứ j từ tài liệu gốc
                splitDocument.addPage(page); // Thêm trang vào tài liệu mới
            }

            // Lưu phần PDF nhỏ vào file
            String splitPdfPath = "D:/2024/Java/Generated/split-" + (i / pagesPerSplit + 1) + ".pdf";
            splitDocument.save(splitPdfPath);
            splitDocument.close(); // Đóng tài liệu PDF đã tách

            // Chuyển đổi phần PDF nhỏ thành DOCX
            String docxOutputPath = "D:/2024/Java/Generated/sample-" + (i / pagesPerSplit + 1) + ".docx";
            convertPdfToDocx(splitPdfPath, docxOutputPath);
            
            // Thêm file DOCX vào danh sách
            docxFiles[i / pagesPerSplit] = docxOutputPath;
        }

        // Đóng tài liệu PDF gốc
        document.close();

        // Nối tất cả các file DOCX lại thành một file duy nhất
        String outputDocx = "D:/2024/Java/Generated/output.docx";
        mergeDocxFiles(docxFiles, outputDocx);

        System.out.println("Tách PDF, chuyển thành DOCX và nối thành công!");
    }

    // Hàm chuyển PDF thành DOCX sử dụng Spire.PDF
    public static void convertPdfToDocx(String pdfFilePath, String docxOutputPath) {
        // Tạo đối tượng PdfDocument từ Spire.PDF
        PdfDocument pdf = new PdfDocument();
        pdf.loadFromFile(pdfFilePath); // Tải PDF vào PdfDocument

        // Chuyển đổi PDF sang DOCX
        pdf.saveToFile(docxOutputPath, FileFormat.DOCX);

        System.out.println("Chuyển đổi PDF sang DOCX thành công: " + docxOutputPath);
    }

    // Hàm nối các file DOCX vào một file duy nhất
    public static void mergeDocxFiles(String[] docxFiles, String outputDocx) throws IOException {
        XWPFDocument mergedDoc = new XWPFDocument();

        for (String docxFile : docxFiles) {
            XWPFDocument doc = new XWPFDocument(new FileInputStream(docxFile));

            List<XWPFParagraph> paragraphs = doc.getParagraphs();
            for (XWPFParagraph paragraph : paragraphs) {
                XWPFParagraph newParagraph = mergedDoc.createParagraph();
                XWPFRun run = newParagraph.createRun();
                run.setText(paragraph.getText()); // Thêm đoạn văn bản vào tài liệu mới
            }
        }

        // Lưu tài liệu đã nối vào file output
        try (FileOutputStream out = new FileOutputStream(outputDocx)) {
            mergedDoc.write(out);
        }

        System.out.println("Nối các file DOCX thành công vào: " + outputDocx);
    }
}
