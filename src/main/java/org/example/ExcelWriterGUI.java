package org.example;

import java.awt.*;
import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileOutputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.swing.*;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriterGUI extends JFrame {

    private TextField databaseField;
    private JTextField usernameFiled;
    private JPasswordField passwordField;
    private JTextArea sqlQueryArea;
    private JTable dataTable;
    private EditableTableModel tableModel;
    private final JButton executeButton = new JButton("Execute Query");
    private final JButton saveButton = new JButton("Save to Excel");

    public ExcelWriterGUI() {
        setTitle("Database to Excel Converter");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new GridLayout());

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        // Connection Panel
        JPanel connectionPanel = new JPanel(new GridBagLayout());
        connectionPanel.setBorder(BorderFactory.createTitledBorder("Database Connection"));
        databaseField = new TextField("jdbc:mysql://localhost:3306/");
        usernameFiled = new JTextField();
        passwordField = new JPasswordField();
        gbc.gridy = 0;
        connectionPanel.add(new JLabel("Database URL: "), gbc);
        gbc.gridy = 1;
        connectionPanel.add(databaseField, gbc);
        gbc.gridy = 2;
        connectionPanel.add(new JLabel("Username: "), gbc);
        gbc.gridy = 3;
        connectionPanel.add(usernameFiled, gbc);
        gbc.gridy = 4;
        connectionPanel.add(new JLabel("Password"), gbc);
        gbc.gridy = 5;
        connectionPanel.add(passwordField, gbc);

        // Query Panel
        JPanel queryPanel = new JPanel(new GridBagLayout());
        queryPanel.setBorder(BorderFactory.createTitledBorder("SQL Query"));
        sqlQueryArea = new JTextArea("SELECT * FROM ", 10, 40);
        sqlQueryArea.setLineWrap(true);
        JScrollPane queryScrollPane = new JScrollPane(sqlQueryArea);
        queryScrollPane.setPreferredSize(new Dimension(400, 150));
        queryScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        gbc.gridx = 0;
        gbc.gridy = 0;
        queryPanel.add(queryScrollPane, gbc);

        // Data Table Panel
        JPanel tablePanel = new JPanel(new BorderLayout());
        tablePanel.setBorder(BorderFactory.createTitledBorder("Data Preview"));

        tableModel = new EditableTableModel();
        dataTable = new JTable();
        dataTable.setModel(tableModel);

        // Set the custom renderer and editor for the header cells
        dataTable.getTableHeader().setDefaultRenderer(new HeaderRenderer());

        JScrollPane tableScrollPane = new JScrollPane(dataTable);
        tableScrollPane.setPreferredSize(new Dimension(600, 200));
        tablePanel.add(tableScrollPane, BorderLayout.CENTER);

        // Save Button Panel
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        buttonPanel.add(executeButton);
        buttonPanel.add(saveButton);

        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.CENTER;
        add(connectionPanel, gbc);

        gbc.gridx = 0;
        gbc.gridy = 1;
        add(queryPanel, gbc);

        gbc.gridx = 0;
        gbc.gridy = 2;
        add(tablePanel, gbc);

        gbc.gridx = 0;
        gbc.gridy = 3;
        add(buttonPanel, gbc);

        executeButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                executeQuery();
            }
        });

        saveButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                saveToExcel();
            }
        });

        displayDataInTable(new ArrayList<>());

        pack();
        setLocationRelativeTo(null);
    }

    private List<List<Object>> fetchDataFromDatabase() throws Exception {
        String databaseUrl = databaseField.getText();
        String username = usernameFiled.getText();
        String password = new String(passwordField.getPassword());
        String sqlQuery = sqlQueryArea.getText();

        try (Connection connection = DriverManager.getConnection(databaseUrl, username, password);
             Statement statement = connection.createStatement();
             ResultSet resultSet = statement.executeQuery(sqlQuery)) {

            ResultSetMetaData metaData = resultSet.getMetaData();
            int numColumns = metaData.getColumnCount();

            List<List<Object>> data = new ArrayList<>();

            // Step 1: Add Column Headers to the Data List
            List<Object> columnHeaders = new ArrayList<>();
            for (int colIndex = 1; colIndex <= numColumns; colIndex++) {
                columnHeaders.add(metaData.getColumnName(colIndex));
            }
            data.add(columnHeaders);

            // Step 2: Add Data Rows to the Data List
            while (resultSet.next()) {
                List<Object> row = new ArrayList<>();
                for (int colIndex = 1; colIndex <= numColumns; colIndex++) {
                    row.add(resultSet.getObject(colIndex));
                }
                data.add(row);
            }
            resultSet.close();
            statement.close();

            return data;
        }
    }

    private void displayDataInTable(List<List<Object>> data) {
        if (data.isEmpty() || data.get(0).isEmpty()) {
            // If there's no data or no column headers, show an empty table
            tableModel.setData(new Object[0][0], new String[0]);
        } else {
            // Convert the List<List<Object>> to a 2D array
            Object[][] tableData = new Object[data.size() - 1][];
            for (int i = 1; i < data.size(); i++) { // Start from 1 to skip the header row
                tableData[i - 1] = data.get(i).toArray();
            }

            // Get the column headers from the first row
            List<Object> originalHeaders = data.get(0);
            List<Object> modifiedHeaders = new ArrayList<>();

            // Capitalize the first letter of each column header and add to modifiedHeaders
            for (Object header : originalHeaders) {
                String headerStr = String.valueOf(header);
                if (!headerStr.isEmpty()) {
                    modifiedHeaders.add(headerStr.substring(0, 1).toUpperCase() + headerStr.substring(1));
                } else {
                    modifiedHeaders.add(headerStr);
                }
            }

            // Update the table model with the new data and modified column headers
            tableModel.setData(tableData, modifiedHeaders.toArray());
        }
    }


    private void executeQuery() {
        try {
            List<List<Object>> data = fetchDataFromDatabase();
            if (data.isEmpty()) {
                JOptionPane.showMessageDialog(this, "No data retrieved from the database.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            displayDataInTable(data);

            JOptionPane.showMessageDialog(this, "Data retrieved from the database successfully!");
        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Error Occurred: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void saveToExcel() {
        try {
            if (tableModel.getRowCount() == 0) {
                JOptionPane.showMessageDialog(this, "No data to save.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            // Prompt user to enter new column headers
            List<Object> newHeaders = new ArrayList<>();
            for (int colIndex = 0; colIndex < tableModel.getColumnCount(); colIndex++) {
                String header = JOptionPane.showInputDialog(this, "Enter new header for column '" + tableModel.getColumnName(colIndex) + "':");
                if (header == null) {
                    // User canceled input
                    return;
                }
                newHeaders.add(header);
            }

            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            int result = fileChooser.showSaveDialog(this);
            if (result == JFileChooser.APPROVE_OPTION) {
                String saveDirectory = fileChooser.getSelectedFile().getAbsolutePath();
                String fileName = JOptionPane.showInputDialog(this, "Enter the Excel file name:", "output.xlsx");

                if (fileName == null || fileName.trim().isEmpty()) {
                    JOptionPane.showMessageDialog(this, "Invalid file name. Please provide a valid name.", "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                Workbook workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("Data");

                // Step 1: Write Column Headers to Excel
                Row headerRow = sheet.createRow(0);
                for (int colIndex = 0; colIndex < tableModel.getColumnCount(); colIndex++) {
                    String header = String.valueOf(newHeaders.get(colIndex));
                    headerRow.createCell(colIndex).setCellValue(header);
                }

                // Step 2: Write Data Rows to Excel
                for (int rowIndex = 0; rowIndex < tableModel.getRowCount(); rowIndex++) {
                    Row excelRow = sheet.createRow(rowIndex + 1); // Start from row 1 to skip header row
                    for (int colIndex = 0; colIndex < tableModel.getColumnCount(); colIndex++) {
                        Cell cell = excelRow.createCell(colIndex);
                        Object value = tableModel.getValueAt(rowIndex, colIndex);
                        if (value instanceof String) {
                            cell.setCellValue((String) value);
                        } else if (value instanceof Date) {
                            cell.setCellValue((Date) value);
                            // Apply date format to the cell
                            CellStyle dateCellStyle = workbook.createCellStyle();
                            CreationHelper creationHelper = workbook.getCreationHelper();
                            dateCellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd"));
                            cell.setCellStyle(dateCellStyle);
                        } else {
                            cell.setCellValue(value.toString());
                        }
                    }
                }

                try (FileOutputStream fileOutputStream = new FileOutputStream(saveDirectory + "/" + fileName)) {
                    workbook.write(fileOutputStream);
                }
                workbook.close();
                JOptionPane.showMessageDialog(this, "Data saved to Excel successfully!");
            }
        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Error Occurred: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }



    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                e.printStackTrace();
            }
            new ExcelWriterGUI().setVisible(true);
        });
    }
}

class EditableTableModel extends AbstractTableModel {

    private Object[][] data = new Object[0][0];
    private String[] columnHeaders = new String[0];

    public void setData(Object[][] data, Object[] columnHeaders) {
        this.data = data;
        this.columnHeaders = new String[columnHeaders.length];
        for (int i = 0; i < columnHeaders.length; i++) {
            this.columnHeaders[i] = String.valueOf(columnHeaders[i]);
        }
        fireTableStructureChanged(); // Use fireTableStructureChanged() instead of fireTableDataChanged()
    }

    @Override
    public int getRowCount() {
        return data.length;
    }

    @Override
    public int getColumnCount() {
        return columnHeaders.length;
    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        return data[rowIndex][columnIndex];
    }

    @Override
    public void setValueAt(Object value, int rowIndex, int columnIndex) {
        data[rowIndex][columnIndex] = value;
        fireTableCellUpdated(rowIndex, columnIndex);
    }

    @Override
    public String getColumnName(int columnIndex) {
        return columnHeaders[columnIndex];
    }

    @Override
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        return true;
    }
}

class HeaderRenderer extends DefaultTableCellRenderer {
    @Override
    public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
        Component component = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
        component.setBackground(Color.LIGHT_GRAY);
        component.setForeground(Color.BLACK);
        return component;
    }
}