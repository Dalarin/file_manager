import javax.swing.*;
import java.awt.*;

public class DocX {

    static JFrame frame = new JFrame("Лабораторная 16");
    JTextArea textArea;
    JScrollPane jScrollPane;
    JTable tableExcel;
    JScrollPane panelForExcel;
    static JPanel panel = new JPanel();
    static JMenuBar menuBar = new JMenuBar();
    static JMenuBar mb = new JMenuBar();
    JMenu fileMenu = new JMenu("Файл");
    JMenuItem createFile = new JMenuItem("Создание файла");
    JMenuItem saveFile = new JMenuItem("Сохранить файл");
    JMenuItem saveFileAs = new JMenuItem("Сохранить как");
    JMenuItem openFile = new JMenuItem("Открыть файл");
    JMenu reports = new JMenu("Просмотр отчётов");
    JMenu admin = new JMenu("Администрирование");
    JMenuItem exit = new JMenuItem("Выход");

    DocX() {
        MenuEngine menuEngine = new MenuEngine(this);
        textArea = new JTextArea();
        textArea.setColumns(55);
        textArea.setRows(28);
        textArea.setLineWrap(true);
        textArea.setWrapStyleWord(true);
        textArea.setEditable(true);
        jScrollPane = new JScrollPane(textArea,
                JScrollPane.VERTICAL_SCROLLBAR_ALWAYS, JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
        jScrollPane.setVisible(false);
        menuBar.add(fileMenu);
        menuBar.add(reports);
        menuBar.add(admin);
        tableExcel = new JTable(150, 65) {
            public boolean getScrollableTracksViewportWidth() {
                return getPreferredSize().width < getParent().getWidth();
            }
        };
        panelForExcel = new JScrollPane(tableExcel);
        panelForExcel.setOpaque(true);
        panelForExcel.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        panelForExcel.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        panelForExcel.setVisible(false);
        menuBar.add(exit);
        fileMenu.add(createFile);
        fileMenu.add(openFile);
        fileMenu.add(saveFile);
        fileMenu.add(saveFileAs);
        createFile.addActionListener(menuEngine);
        openFile.addActionListener(menuEngine);
        saveFile.addActionListener(menuEngine);
        saveFileAs.addActionListener(menuEngine);
        exit.addActionListener(menuEngine);
        panel.add(panelForExcel);
        panel.add("Center", jScrollPane);
    }


    private void run() {
        panel.setVisible(true);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        UIManager.put("MenuBar.background", Color.lightGray);// цвет для верхней панели
        frame.setJMenuBar(mb);
        frame.setJMenuBar(menuBar);
        frame.setSize(620, 550);
        frame.setResizable(false);
        frame.setVisible(true);
        frame.setLayout(new BorderLayout());
        frame.add(panel, BorderLayout.CENTER);
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
    }

    public static void main(final String[] args) {
        DocX DocX = new DocX();
        DocX.run();

    }
}
// h
