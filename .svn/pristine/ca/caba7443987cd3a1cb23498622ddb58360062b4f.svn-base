package com.otc.selenium;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JList;
import javax.swing.JPanel;
import javax.swing.JScrollPane;

public class Driver extends JPanel implements ActionListener {

    private static final long serialVersionUID = 1L;

    String exeresultprefix, exeresultsuffix, errorsnapprefix, errorsnapsuffix, errorsnapshotforpass, errorsnapshotforfail, browser;

    public static JFrame frame;
    JButton addFile, startExecution, remove, browsertype;
    DefaultListModel<File> listModel;
    JList<File> log;
    JFileChooser fc;
    int i;
    static String key;
    public int fileCounter;
    public static GenericMethods call11;
    public static File[] InputExcelPath;
    public static File[] InputExcelPathTemp = new File[10];
    public static int tempLength;
    public static int tempCount;

    public Driver() {

        super(new BorderLayout());

        // Setting the properties of List
        listModel = new DefaultListModel<File>();
        log = new JList<File>(listModel);

        JScrollPane logScrollPane = new JScrollPane(log);

        // Creating a file chooser
        fc = new JFileChooser();
        fc.setMultiSelectionEnabled(true);

        // Creating the Add File button
        addFile = new JButton("Add File");
        addFile.addActionListener(this);

        // Creating the save button
        startExecution = new JButton("Start Execution");
        startExecution.addActionListener(this);

        // Creating the remove button
        remove = new JButton("Remove");
        remove.addActionListener(this);

        // Layout buttons in a separate panel
        JPanel buttonPanel = new JPanel();
        buttonPanel.add(addFile);
        buttonPanel.add(startExecution);
        buttonPanel.add(remove);

        // Adding the buttons and the log to this panel.
        add(buttonPanel, BorderLayout.PAGE_START);
        add(logScrollPane, BorderLayout.CENTER);
    }

    public void actionPerformed(ActionEvent e) {

        // Handling Add File button action.
        if (e.getSource() == addFile) {

            try {

                if (fc.showOpenDialog(Driver.this) == JFileChooser.APPROVE_OPTION) {

                    InputExcelPath = fc.getSelectedFiles();

                    for (i = 0; i < InputExcelPath.length; i++) {

                        listModel.addElement(InputExcelPath[i]);
                        InputExcelPathTemp[tempLength + i] = InputExcelPath[i];

                        tempCount++;
                    }

                    tempLength = tempCount;

                }
            } catch (Exception c) {
                // System.out.println(c.getMessage());
            }
        }

        // Handling Start Execution button action.
        else if (e.getSource() == startExecution) {

            int end = tempLength;

            for (int i = 0; i < end; i++) {
                for (int j = i + 1; j < end; j++) {
                    if (InputExcelPathTemp[i] == InputExcelPathTemp[j]) {
                        int shiftLeft = j;
                        for (int k = j + 1; k < end; k++, shiftLeft++) {
                            InputExcelPathTemp[shiftLeft] = InputExcelPathTemp[k];
                        }
                        end--;
                        j--;
                    }
                }
            }

            File[] whitelist = new File[end];
            for (int i = 0; i < end; i++) {
                whitelist[i] = InputExcelPathTemp[i];
            }

            frame.dispose();

            try {
                if (whitelist != null && call11 != null) {
                    call11.ExcelRead(whitelist);
                }

            }

            catch (Exception e1) {

                System.out.println("error in calling excelread" + e1.getMessage());
                e1.printStackTrace();

            }

        }

        else if (e.getSource() == remove) {

            try {

                listModel = (DefaultListModel<File>) log.getModel();
                int selectedIndex = log.getSelectedIndex();

                if (selectedIndex >= 0) {
                    listModel.remove(selectedIndex);
                }

                for (int k = 0; k < InputExcelPathTemp.length - 1; k++) {
                    if (k >= selectedIndex) {
                        InputExcelPathTemp[k] = InputExcelPathTemp[k + 1];

                    }
                }
            } catch (Exception ex) {
                ex.printStackTrace();

            }
        }
    }

    // Creating GUI window
    private static void createAndShowGUI() {

        frame = new JFrame("Excel files to be Tested");
        frame.setPreferredSize(new Dimension(500, 200));

        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setBounds(400, 400, 100, 100);

        JComponent newContentPane = new Driver();
        newContentPane.setOpaque(true); // content panes must be opaque
        frame.setContentPane(newContentPane);

        frame.pack();
        frame.setVisible(true);

    }

    public static void main(String[] args) {
        // Initializing th driver.
        Driver d = new Driver();
        call11 = new GenericMethods(d);
        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });

    }
}
