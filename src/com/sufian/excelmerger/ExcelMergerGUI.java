package com.sufian.excelmerger;

import javax.swing.*;

import com.sufian.excelmerger.ExcelTransactionExtractor.Transaction;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelMergerGUI extends JFrame {
    /**
	 * 
	 */
	private static final long serialVersionUID = -2890840092653497187L;
	private JLabel monthLabel, inputPathLabel, outputPathLabel;
    private JComboBox<String> monthComboBox;
    private JTextField inputPathField, outputPathField;
    private JButton chooseInputButton, chooseOutputButton, generateButton;

    public ExcelMergerGUI() {
        // Set system look and feel
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }

        setTitle("Excel Merger");
        setSize(400, 200);
        setLocationRelativeTo(null); // Center the frame on the screen
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // Create components
        monthLabel = new JLabel("Month:");
        inputPathLabel = new JLabel("Input Path:");
        outputPathLabel = new JLabel("Output Path:");

        String[] months = {"January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December"};
        monthComboBox = new JComboBox<>(months);

        inputPathField = new JTextField(20); // Removed size specification
        outputPathField = new JTextField(20); // Removed size specification
        inputPathField.setEditable(false);
        outputPathField.setEditable(false);

        chooseInputButton = new JButton("Choose");
        chooseOutputButton = new JButton("Choose");
        generateButton = new JButton("Generate");

        // Add action listeners
        chooseInputButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                chooseInputPath();
            }
        });

        chooseOutputButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                chooseOutputPath();
            }
        });

        generateButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                generateExcel();
            }
        });

        // Set layout
        setLayout(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.WEST;

        // Add components to the frame
        add(monthLabel, gbc);
        gbc.gridy++;
        add(inputPathLabel, gbc);
        gbc.gridy++;
        add(outputPathLabel, gbc);

        gbc.gridx = 1;
        gbc.gridy = 0;
        add(monthComboBox, gbc);
        gbc.gridy++;
        gbc.fill = GridBagConstraints.HORIZONTAL; // Allow horizontal expansion
        add(inputPathField, gbc);
        gbc.gridy++;
        add(outputPathField, gbc);
        gbc.gridy++;
        add(generateButton, gbc);

        gbc.gridx = 2;
        gbc.gridy = 1;
        gbc.fill = GridBagConstraints.NONE; // Reset fill
        add(chooseInputButton, gbc);
        gbc.gridy++;
        add(chooseOutputButton, gbc);
        

        setVisible(true);
    }

    private void chooseInputPath() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        String defaultPath = "";
        if(inputPathField.getText().isEmpty() || inputPathField.getText() == null)
        	defaultPath = "C:\\";
        else 
        	defaultPath = inputPathField.getText();
        fileChooser.setCurrentDirectory(new File(defaultPath));
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            inputPathField.setText(fileChooser.getSelectedFile().getAbsolutePath());
        }
        if(outputPathField.getText() == null || outputPathField.getText().isEmpty()) {
        	outputPathField.setText(fileChooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void chooseOutputPath() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        String defaultPath = "";
        if(outputPathField.getText().isEmpty() || outputPathField.getText() == null)
        	defaultPath = "C:\\";
        else 
        	defaultPath = outputPathField.getText();
        fileChooser.setCurrentDirectory(new File(defaultPath));
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            outputPathField.setText(fileChooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void generateExcel() {
        // Implement logic to generate Excel file
        String month = monthComboBox.getSelectedItem().toString();
        String inputPath = inputPathField.getText();
        String outputPath = outputPathField.getText();
        String filePath = outputPath + "\\" + month + ".xlsx";
        try {
        	if(!month.isEmpty() && !inputPath.isEmpty() && !outputPath.isEmpty()) {
        		List<Transaction> allTransactions = new ArrayList<>();
        		allTransactions = ExcelTransactionExtractor.extractFromAllExcel(inputPath, month);
        		if(allTransactions.size() == 0) {
        			JOptionPane.showMessageDialog(null, "There are no transactions for the month of " + month);
        		} else {
        			ExcelCreator.createExcelFile(allTransactions, filePath);
            		JOptionPane.showMessageDialog(null, "Excel file successfully created");
        		}
        	}
        	else
        		JOptionPane.showMessageDialog(null, "Please select valid directory for input and output path");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        // and input/output paths from inputPathField.getText() and outputPathField.getText()
    }

    public static void main(String[] args) {
        new ExcelMergerGUI();
    }
}