import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Main {
    private JFrame frame;
    private JButton openButton;

    public Main() {
        initialize();
    }

    private void initialize() {
        frame = new JFrame("Excel Reader");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(600, 400);

        JPanel panel = new JPanel();
        openButton = new JButton("Open Excel File");
        openButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                int result = fileChooser.showOpenDialog(frame);
                if (result == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    readExcelFile(selectedFile);
                }
            }
        });

        panel.add(openButton);
        frame.add(panel);
        frame.setVisible(true);
    }

    private void readExcelFile(File file) {
        try (java.io.InputStream inputStream = new java.io.FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(inputStream)) {

            JOptionPane.showMessageDialog(frame, "Excel file loaded successfully!");
            System.out.println("Excel file loaded successfully!");

            // Create threads for processing each sheet
            Thread threadOld = new Thread(() -> processSheet(workbook, "MainOld"));
            Thread threadNew = new Thread(() -> processSheet(workbook, "MainNew"));

            threadOld.start();
            threadNew.start();

            threadOld.join();
            threadNew.join();

            // Analyze differences
            System.out.println("Imbalance Accounts:");
            accountDataMap.forEach((voiceId, data) -> {
                double debitDiff = data.debitOld - data.debitNew;
                double creditDiff = data.creditOld - data.creditNew;
                
                if (debitDiff != 0.0 || creditDiff != 0.0) {
                    System.out.printf("VoiceID: %s | Debit Diff: %.2f | Credit Diff: %.2f%n", 
                                     voiceId, debitDiff, creditDiff);
                }
            });

        } catch (Exception e) {
            JOptionPane.showMessageDialog(frame, "Error reading Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void processSheet(Workbook workbook, String sheetName) {
        try {
            Sheet sheet = workbook.getSheet(sheetName);
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row != null) {
                    Cell voiceIdCell = row.getCell(0);
                    Cell debitCell = row.getCell(3);
                    Cell creditCell = row.getCell(4);

                    String voiceId = voiceIdCell.getStringCellValue();
                    double debit = debitCell != null ? debitCell.getNumericCellValue() : 0.0;
                    double credit = creditCell != null ? creditCell.getNumericCellValue() : 0.0;

                    accountDataMap.compute(voiceId, (key, existing) -> {
                        if (existing == null) {
                            return new AccountData(debit, 0.0, credit, 0.0);
                        } else {
                            if (sheetName.equals("MainOld")) {
                                return new AccountData(debit + existing.debitOld, existing.debitNew, 
                                                      credit + existing.creditOld, existing.creditNew);
                            } else {
                                return new AccountData(existing.debitOld, debit + existing.debitNew, 
                                                      existing.creditOld, credit + existing.creditNew);
                            }
                        }
                    });
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static class AccountData {
        double debitOld = 0.0;
        double debitNew = 0.0;
        double creditOld = 0.0;
        double creditNew = 0.0;

        public AccountData(double debitOld, double debitNew, double creditOld, double creditNew) {
            this.debitOld = debitOld;
            this.debitNew = debitNew;
            this.creditOld = creditOld;
            this.creditNew = creditNew;
        }

        public double getDebitOld() { return debitOld; }
        public double getDebitNew() { return debitNew; }
        public double getCreditOld() { return creditOld; }
        public double getCreditNew() { return creditNew; }
    }

    private static final java.util.concurrent.ConcurrentHashMap<String, AccountData> accountDataMap = new java.util.concurrent.ConcurrentHashMap<>();

    public static void main(String[] args) {
        EventQueue.invokeLater(new Runnable() {
            @Override
            public void run() {
                try {
                    new Main();
                } catch (Throwable e) {
                    e.printStackTrace();
                }
            }
        });
    }
}
