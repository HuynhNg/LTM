package org.sample.BTCK;

import com.spire.doc.Document;
import com.spire.doc.FileFormat; // Sử dụng FileFormat từ Spire.Doc
import com.spire.doc.Section;
import com.spire.pdf.PdfDocument;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;

import java.io.File;
import java.io.IOException;

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
        pdf.saveToFile(docxOutputPath, com.spire.pdf.FileFormat.DOCX);

        System.out.println("Chuyển đổi PDF sang DOCX thành công: " + docxOutputPath);
    }

    // Hàm nối các file DOCX vào một file duy nhất sử dụng Spire.Doc
    public static void mergeDocxFiles(String[] docxFiles, String outputDocx) throws IOException {
        // Tạo một tài liệu Word mới
        Document mergedDoc = new Document();

        // Duyệt qua từng file DOCX trong danh sách
        for (String docxFile : docxFiles) {
            // Mở tài liệu DOCX từ file
            Document doc = new Document();
            doc.loadFromFile(docxFile);

            // Duyệt qua tất cả các phần trong tài liệu
            for (int j = 0; j < doc.getSections().getCount(); j++) {
                Section section = doc.getSections().get(j);
                // Thêm phần vào tài liệu chính
                mergedDoc.importSection(section);
            }
        }// Lưu tài liệu đã nối vào file output
        mergedDoc.saveToFile(outputDocx, FileFormat.Docx);

        System.out.println("Nối các file DOCX thành công vào: " + outputDocx);
    }
}
