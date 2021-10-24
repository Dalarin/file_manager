import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;


public class MenuEngine implements ActionListener {
    DocX parent;
    workWithFile fileAct;

    public MenuEngine(DocX parent) {
        this.parent = parent;
        this.fileAct = new workWithFile(parent);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        Object src = e.getSource();
        if (src instanceof JMenuItem) {
            src = (JMenuItem) e.getSource();
            try {
                if (parent.createFile.equals(src)) {
                    fileAct.getPathFile(1);
                } else if (parent.openFile.equals(src)) {
                    fileAct.getPathFile(2);
                } else if (parent.saveFile.equals(src)) {
                    if (fileAct.getFileFormat() != null) fileAct.getPathFile(3);
                } else if (parent.saveFileAs.equals(src)) {
                    if (fileAct.getFileFormat() != null) fileAct.getPathFile(4);
                } else if (parent.exit.equals(src)) {
                    DocX.frame.dispose();
                }
            } catch (IOException | InvalidFormatException ex) {
                JOptionPane.showMessageDialog(null, "Ошибка в работе программы", "Ошибка", JOptionPane.ERROR_MESSAGE);
                System.out.println(ex);
            }
        }
    }
}
// h


