
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.InvalidPathException;
import java.nio.file.Paths;
import java.util.*;
import java.util.List;

public class workWithFile {
    DocX parent;
    private String fileFormat;
    private File file;

    workWithFile(DocX parent) {
        this.parent = parent;

    }

    private int getNumColumns(Sheet sheet, int rowN) {
        Row row = sheet.getRow(rowN);
        return (int) (row == null ? 0 : row.getLastCellNum());
    }

    private Object[][] getTableData() {
        DefaultTableModel tableModel = (DefaultTableModel) parent.tableExcel.getModel();
        tableModel.fireTableDataChanged();
        parent.tableExcel.repaint();
        DocX.panel.updateUI();
        int nRow = tableModel.getRowCount();
        int nCol = tableModel.getColumnCount();
        Object[][] tableData = new Object[nRow][nCol];
        for (int i = 0; i < nRow; i++) {
            for (int j = 0; j < nCol; j++) {
                tableData[i][j] = tableModel.getValueAt(i, j);
            }
        }
        return tableData;
    }

    public String getFileFormat() {
        return fileFormat;
    }

    private void writeDOCXFile() {
        try {
            XWPFDocument document = new XWPFDocument();
            XWPFParagraph tmpParagraph = document.createParagraph();
            XWPFRun tmpRun = tmpParagraph.createRun();
            String[] paragraph = parent.textArea.getText().split("\\n");
            if (parent.textArea.getText().contains("\n")) {
                for (int i = 0; i < paragraph.length; i++) {
                    tmpRun.addBreak();
                    tmpRun.setText(paragraph[i], i);
                }
            } else {
                tmpRun.setText(parent.textArea.getText());
            }
            document.write(new FileOutputStream(String.valueOf(this.file)));
            errorMessage("Сохранен", "Файл успешно сохранен", 1);
        } catch (IOException ignored) {
            errorMessage("Ошибка", "Файл не сохранен", 0);
            System.out.println(ignored);
        }
    }

    private void writeTXTFile() {
        try {
            FileWriter fileWriter = new FileWriter(this.file);
            PrintWriter printWriter = new PrintWriter(fileWriter);
            printWriter.printf(parent.textArea.getText());
            printWriter.close();
            errorMessage("Сохранен", "Файл успешно сохранен", 1);

        } catch (Exception e) {
            errorMessage("Ошибка", "Файл не сохранен", 0);

        }
    }

    private void writeXLSXFile() throws IOException {
        FileOutputStream file = new FileOutputStream(String.valueOf(this.file));
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.getNumberOfSheets() != 0 ? workbook.getSheetAt(0) : workbook.createSheet("name");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        int rows = parent.tableExcel.getRowCount();
        int columns = parent.tableExcel.getColumnCount();
        Object[][] data = getTableData();
        for (int i = 0; i < rows; i++) {
            row = sheet.createRow(i);
            for (int j = 0; j < columns; j++) {
                cell = row.createCell(j);
                String value = data[i][j] != null ? data[i][j].toString() : " ";
                cell.setCellValue(value);
            }
        }
        workbook.write(file);
        workbook.close();
        errorMessage("Успех", "Файл успешно сохранен", 1);
    }

    private void writeCSVFile() {
        try {
            PrintWriter printWriter = new PrintWriter(new File(String.valueOf(this.file)), "Cp1251");
            Object[][] data = getTableData();
            StringBuilder builder = new StringBuilder();
            for (Object[] datum : data) {
                for (Object o : datum) {
                    String value = o != null ? o + ", " : " ";
                    builder.append(value);
                }
                builder.append('\n');
            }
            printWriter.printf(String.valueOf(builder));
            printWriter.close();
            errorMessage("Сохранен", "Файл успешно сохранен", 1);
        } catch (IOException ex) {
            errorMessage("Ошибка", "Ошибка в сохранении файла", 0);
        }
    }

    private void saveFile() throws IOException, InvalidFormatException {
        switch (fileFormat) {
            case "docx" -> writeDOCXFile();
            case "txt" -> writeTXTFile();
            case "xlsx" -> writeXLSXFile();
            case "csv" -> writeCSVFile();
            default -> errorMessage("Ошибка", "Неожиданный формат: " + fileFormat, 0);
        }
    }

    private void createFile(String pathOfFile) throws IOException {
        File newFile = new File(pathOfFile);
        newFile.getParentFile().mkdirs();
        if (!newFile.exists() && newFile.getName().length() != 0) {
            try {
                String[] parts = newFile.getName().split("\\.");
                fileFormat = parts[parts.length - 1];
                Paths.get(pathOfFile);
                file = new File(String.valueOf(newFile));
                newFile.createNewFile();
                if (parts[parts.length - 1].equals("txt") || parts[parts.length - 1].equals("docx")) {
                    parent.jScrollPane.setVisible(true);
                    DocX.panel.updateUI();
                } else if (parts[parts.length - 1].equals("xlsx") || parts[parts.length - 1].equals("csv")) {
                    parent.panelForExcel.setVisible(true);
                    DocX.panel.updateUI();
                }
            } catch (InvalidPathException ex) {
                errorMessage("Ошибка", "Некорректное название файла", 0);
            }
        } else {
            errorMessage("Ошибка", "Файл с таким названием уже существует", 0);
        }

    }

    private void inputNameWindow(String pathofFile, Integer typeOfOperation) {
        String[] fileFormat = {"txt", "docx", "xlsx", "csv"};
        JButton ok = new JButton("OK");
        JPanel panel = new JPanel();
        JTextField inputName = new JTextField();
        JComboBox<String> cb = new JComboBox<>(fileFormat);
        JFrame frame = new JFrame("Название файла");
        frame.setSize(300, 100);
        frame.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
        frame.setResizable(false);
        frame.setVisible(true);
        frame.setLayout(new BorderLayout());
        frame.add(panel, BorderLayout.CENTER);
        inputName.setPreferredSize(new Dimension(100, 30));
        if (typeOfOperation == 1) panel.add(cb);
        panel.add(inputName, "south");
        panel.add(ok);
        String file = this.fileFormat;
        ok.addActionListener(e -> {
            String nameFile = typeOfOperation == 4 ? inputName.getText() + '.' + file : inputName.getText() + '.' + Objects.requireNonNull(cb.getSelectedItem());
            try {
                createFile(pathofFile + '\\' + nameFile);
                if (typeOfOperation == 4) saveFile();
                frame.dispose();

            } catch (IOException | InvalidFormatException ex) {
                ex.printStackTrace();
            }
        });

    }

    private void errorMessage(String title, String message, Integer code) {
        JOptionPane.showMessageDialog(null, message, title, code);
    }

    private void parseDOCXFile() { // парсим DOCX файл с помощью библиотеки Apache
        parent.jScrollPane.setVisible(true);
        DocX.panel.updateUI();
        try {
            StringBuilder textInFile = new StringBuilder();
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            XWPFDocument document = new XWPFDocument(fis);
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (XWPFParagraph para : paragraphs) {
                textInFile.append(para.getText());
            }
            parent.textArea.setText(textInFile.toString());
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void parseXLSXFile() {
        //parent.panelForExcel.setVisible(true);
        try {
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0);
            int numberRows = sheet.getLastRowNum();
            DefaultTableModel model = (DefaultTableModel) parent.tableExcel.getModel();
            for (int i = 0; i < numberRows; i++) {
                Row row = sheet.getRow(i);
                for (int j = 0; j < getNumColumns(sheet, i); j++) {
                    Cell cell = row.getCell(j);
                    String value = cell != null ? cell.toString() : " ";
                    model.setValueAt(value, i, j);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        parent.panelForExcel.setVisible(true);
        DocX.panel.updateUI();
    }

    private void parseTXTFile() throws IOException {
        parent.jScrollPane.setVisible(true);
        DocX.panel.updateUI();
        BufferedReader reader = new BufferedReader(new FileReader(file));
        StringBuilder stringBuilder = new StringBuilder();
        String line;
        String ls = System.getProperty("line.separator");
        while ((line = reader.readLine()) != null) {
            stringBuilder.append(line);
            stringBuilder.append(ls);
        }
        stringBuilder.deleteCharAt(stringBuilder.length() - 1);
        reader.close();
        parent.textArea.setText(String.valueOf(stringBuilder));
    }

    private void parseCSVFile() throws IOException {
        parent.panelForExcel.setVisible(true);
        ArrayList<ArrayList<String>> data = new ArrayList<>();
        DefaultTableModel model = (DefaultTableModel) parent.tableExcel.getModel();
        try (BufferedReader br = new BufferedReader(new FileReader(file))) {
            String line;
            String ls = System.getProperty("line.separator");
            while ((line = br.readLine()) != null) {
                ArrayList<String> subData = new ArrayList<>(Arrays.asList(line.split("\\,")));
                data.add(subData);
            }
            for (int i = 0; i < data.size(); i++) {
                for (int j = 0; j < data.get(i).size(); j++) {
                    model.setValueAt(data.get(i).get(j), i, j);
                }
            }
        }
        DocX.panel.updateUI();
    }

    private void openFile(JFileChooser fileChooser) throws IOException {
        parent.panelForExcel.setVisible(false);
        parent.jScrollPane.setVisible(false);
        DocX.panel.updateUI();
        fileChooser.setDialogTitle("Выбор файла");
        FileNameExtensionFilter filter = new FileNameExtensionFilter("TXT, Word, Excel, CSV", "txt", "docx", "xlsx", "csv");
        fileChooser.setFileFilter(filter);
        if (fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
            String[] parts = String.valueOf(fileChooser.getSelectedFile()).split("\\.");
            file = new File(String.valueOf(fileChooser.getSelectedFile()));
            fileFormat = parts[parts.length - 1];
            switch (parts[parts.length - 1]) {
                case "xlsx" -> parseXLSXFile();
                case "docx" -> parseDOCXFile();
                case "txt" -> parseTXTFile();
                case "csv" -> parseCSVFile();
                default -> errorMessage("Ошибка", "Неожиданный формат: " + parts[parts.length - 1], 0);
            }

        }
    }

    private void chooseDirectory(JFileChooser fileChooser, Integer typeOfOperation) {
        parent.panelForExcel.setVisible(false);
        parent.jScrollPane.setVisible(false);
        DocX.panel.updateUI();
        fileChooser.setCurrentDirectory(new java.io.File("."));
        fileChooser.setDialogTitle("Выбор директории файла");
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        fileChooser.setAcceptAllFileFilterUsed(false);
        if (fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
            inputNameWindow(fileChooser.getSelectedFile().getAbsolutePath(), typeOfOperation);
        }
    }

    public void getPathFile(Integer typeofOperation) throws IOException, InvalidFormatException { // получаем ссылку на файл
        JFileChooser fileChooser = new JFileChooser();
        switch (typeofOperation) {
            case 1, 4 -> chooseDirectory(fileChooser, typeofOperation);
            case 2 -> openFile(fileChooser);
            case 3 -> saveFile();
        }
    }
}
// h // h
