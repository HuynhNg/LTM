package org.sample.BTCK;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import com.spire.pdf.PdfDocument;
import com.spire.pdf.FileFormat;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;

import java.io.*;
import java.util.List;

public class App {
    public static void main(String[] args) throws IOException, InvalidFormatException {
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

    // Hàm nối các file DOCX vào một file duy nhất mà không mất định dạng
    public static void mergeDocxFiles(String[] docxFiles, String outputDocx) throws IOException, InvalidFormatException {
        XWPFDocument mergedDoc = new XWPFDocument();

        for (String docxFile : docxFiles) {
            XWPFDocument doc = new XWPFDocument(new FileInputStream(docxFile));

            // Sao chép tất cả các đoạn văn từ file DOCX vào tài liệu mergedDoc
            copyParagraphs(doc, mergedDoc);

            // Sao chép các bảng từ file DOCX vào tài liệu mergedDoc
            copyTables(doc, mergedDoc);

            // Sao chép hình ảnh (Pictures) từ tài liệu nguồn vào tài liệu kết quả
            copyPictures(doc, mergedDoc);
        }

        // Lưu tài liệu đã nối vào file output
        try (FileOutputStream out = new FileOutputStream(outputDocx)) {
            mergedDoc.write(out);
        }

        System.out.println("Nối các file DOCX thành công vào: " + outputDocx);
    }

    // Sao chép các đoạn văn bản từ tài liệu nguồn sang tài liệu đích
    public static void copyParagraphs(XWPFDocument srcDoc, XWPFDocument destDoc) {
        for (XWPFParagraph paragraph : srcDoc.getParagraphs()) {
            XWPFParagraph newParagraph = destDoc.createParagraph();
            for (XWPFRun run : paragraph.getRuns()) {
                XWPFRun newRun = newParagraph.createRun();
                newRun.setText(run.toString()); // Sao chép nội dung văn bản

                // Sao chép định dạng của đoạn văn
                newRun.setBold(run.isBold());
                newRun.setItalic(run.isItalic());
                newRun.setStrike(run.isStrikeThrough());
                newRun.setUnderline(run.getUnderline());
                newRun.setFontSize(run.getFontSize());
                newRun.setFontFamily(run.getFontFamily());
                newRun.setColor(run.getColor());
            }

            // Sao chép các thuộc tính của đoạn văn (căn lề, v.v.)
            newParagraph.setAlignment(paragraph.getAlignment());
            newParagraph.setVerticalAlignment(paragraph.getVerticalAlignment());
        }
    }

    // Sao chép các bảng từ tài liệu nguồn sang tài liệu đích
    public static void copyTables(XWPFDocument srcDoc, XWPFDocument destDoc) {
        for (XWPFTable table : srcDoc.getTables()) {
            XWPFTable newTable = destDoc.createTable();

            // Sao chép từng dòng trong bảng
            for (XWPFTableRow row : table.getRows()) {
                XWPFTableRow newRow = newTable.createRow();
                for (XWPFTableCell cell : row.getTableCells()) {
                    XWPFTableCell newCell = newRow.createCell();
                    newCell.setText(cell.getText());

                    // Sao chép thêm các định dạng của bảng (nếu cần)
                }
            }
        }
    }

    public static void copyPictures(XWPFDocument srcDoc, XWPFDocument destDoc) throws IOException, InvalidFormatException {
        for (XWPFPictureData picture : srcDoc.getAllPictures()) {
            byte[] pictureData = picture.getData();
            int pictureType = getPictureType(picture.suggestFileExtension());

            // Thêm hình ảnh vào tài liệu đích (DOCX)
            destDoc.addPictureData(pictureData, pictureType);

            // Tạo một đoạn văn mới và chèn hình ảnh vào đó
            XWPFParagraph paragraph = destDoc.createParagraph();
            XWPFRun run = paragraph.createRun();

            try (ByteArrayInputStream bais = new ByteArrayInputStream(pictureData)) {
                // Chèn hình ảnh vào đoạn văn (cung cấp chiều rộng và chiều cao)
                run.addPicture(bais, pictureType, picture.getFileName(), Units.toEMU(200), Units.toEMU(200)); // Kích thước có thể thay đổi
            } catch (Exception e) {
                System.err.println("Lỗi khi chèn hình ảnh: " + e.getMessage());
            }
        }
    }

    // Phương thức hỗ trợ để chuyển phần mở rộng tệp sang kiểu hình ảnh
    private static int getPictureType(String fileExtension) {
        switch (fileExtension.toLowerCase()) {
            case "jpeg":
            case "jpg":
                return XWPFDocument.PICTURE_TYPE_JPEG;
            case "png":
                return XWPFDocument.PICTURE_TYPE_PNG;
            case "gif":
                return XWPFDocument.PICTURE_TYPE_GIF;
            case "bmp":
                return XWPFDocument.PICTURE_TYPE_BMP;
            default:
                throw new IllegalArgumentException("Không hỗ trợ định dạng hình ảnh: " + fileExtension);
        }
    }



}
