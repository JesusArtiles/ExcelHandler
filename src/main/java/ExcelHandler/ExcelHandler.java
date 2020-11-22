package ExcelHandler;

import Model.Volunteer;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

public class ExcelHandler {
    private FileInputStream fis;
    private FileOutputStream fos;
    private XSSFSheet sh;
    private XSSFWorkbook wb;
    private List<Volunteer> volunteerList;
    private String root;

    public ExcelHandler(List<Volunteer> volunteerList, String root){
        this.volunteerList = volunteerList;
        this.root = root;
    }

    public void initiate() throws Exception {
        File excel = new File(root);
        fis = new FileInputStream(excel);
        wb = new XSSFWorkbook(fis);
        sh = wb.getSheetAt(0);
    }

    public void addVolunteer(Volunteer v) throws IOException {
        volunteerList.add(v);
        writeVolunteersInExcel();
    }

    public void deleteVolunteer(Volunteer v) throws IOException {
        Iterator<Volunteer> iter = volunteerList.iterator();

        while (iter.hasNext()) {
            Volunteer volunteer = iter.next();

            if (volunteer.getEmail().equals(v.getEmail()))
                iter.remove();
        }
        clearExcel();
        writeVolunteersInExcel();
    }

    public void clearExcel() throws IOException {
        int counter = 0;
            for(Row r:sh) {
                if(r != null) {
                    r.removeCell(r.getCell(0));
                    r.removeCell(r.getCell(1));
                    counter++;
                }
            }
            for(int i = 0;i < counter;i++){
                sh.removeRow(sh.getRow(i));
            }
            performChangesInExcel();
    }

    public void writeVolunteersInExcel() throws IOException {
        int counter = 1;

        sh.createRow(0).createCell(0).setCellValue("NAME");
        sh.getRow(0).createCell(1).setCellValue("EMAIL");

        for (Volunteer v:volunteerList){
            sh.createRow(counter).createCell(0).setCellValue(v.getName());
            sh.getRow(counter).createCell(1).setCellValue(v.getEmail());
            counter++;
        }
        performChangesInExcel();
    }

    public String[] readExcel(){
        String[] content = new String[sh.getLastRowNum()+1];
        int counter = 0;
        for(Row r:sh) {
            if(r != null) {
                String name = r.getCell(0).toString();
                String email = r.getCell(1).toString();
                content[counter] = name+" "+email;
                counter++;
            }
        }
        return content;
    }

    public void downloadExcel(String destination){
        File source = new File("./testdata.xlsx");
        File dest = new File(destination);
        try {
            FileUtils.copyFile(source, dest);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void performChangesInExcel() throws IOException {
        File excel = new File(root);
        fos = new FileOutputStream(excel);
        wb.write(fos);
    }




}
