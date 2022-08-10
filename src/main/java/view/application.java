package view;

import Util.WordtoPdfUtil;

import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public class application extends JFrame {
    public  void windows() {
        JFrame jf =new JFrame("WORD 转 PDF");
        jf.setVisible(true);
        jf.setSize(300,300);
        jf.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        JButton jb = new JButton("选择文件");
        jb.setBounds(100,100,100,100);
        jb.setBackground(Color.orange);
        OpenListener ol = new OpenListener();
        jb.addActionListener(ol);
        jf.add(jb);
    }
    public static void main(String[] args) {
        new application().windows();
    }
}

class OpenListener implements ActionListener{
    @Override
    public void actionPerformed(ActionEvent e) {
        JFileChooser fc = new JFileChooser();
        fc.addChoosableFileFilter(new FileFilter() {
            @Override
            public boolean accept(File f) {
                if (f.getName().toLowerCase(Locale.ROOT).endsWith(".doc")||
                        f.getName().toLowerCase(Locale.ROOT).endsWith(".docx")){
                    return true;
                }
                return false;
            }
            @Override
            public String getDescription() {
                return "World文件(*.docx/*.doc)";
            }
        });
        fc.setMultiSelectionEnabled(true);
        fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        fc.showOpenDialog(null);
        File[] files = fc.getSelectedFiles();
        List<File> fileList = new ArrayList<>();
        WordtoPdfUtil wordtoPdfUtil = new WordtoPdfUtil();
        Integer index = 0,count = 0;

        try{
            for (int i = 0; i < files.length; i++) {
                if (files[i].isDirectory()){
                    fileList = WordtoPdfUtil.getFileList(files[i].getPath());
                    for (File f:fileList) {
                        index++;
                        wordtoPdfUtil.WordToPdf(f.getPath());
                    }
                }else{
                    count++;index++;
                    wordtoPdfUtil.WordToPdf(files[i].getPath());
                }
            }
        } catch (Exception exception) {
            JOptionPane.showMessageDialog(null, "转换失败", "", JOptionPane.ERROR_MESSAGE);
        }
        if (index == count+fileList.size()&&index!=0){
            JOptionPane.showMessageDialog(null,"转换成功", "", JOptionPane.INFORMATION_MESSAGE);
        }

    }
}
