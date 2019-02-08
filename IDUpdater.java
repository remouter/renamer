//javac -cp poi-3.9-20121203.jar IDUpdater.java
//jar cfm IDUpdater.jar MANIFEST.MF IDUpdater.class poi-3.9-20121203.jar

//java -cp .:/home/exp.exactpro.com/oleg.legkov/Desktop/Corrector_3.0/lib/poi-3.9-20121203.jar IDUpdater
//javac -cp .:/home/exp.exactpro.com/oleg.legkov/Desktop/Corrector_3.0/lib/poi-3.9-20121203.jar IDUpdater.java

//jar cfm IDUpdater.jar MANIFEST.MF IDUpdater.class /home/exp.exactpro.com/oleg.legkov/Desktop/poi-3.9-20121203.jar
//java -jar IDUpdater.jar -cp .:/home/exp.exactpro.com/oleg.legkov/Desktop/Corrector_3.0/lib/poi-3.9-20121203.jar


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import java.io.*;
import java.util.*;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.*;

public class IDUpdater{
	private FileInputStream in;
	private FileOutputStream out;
	private File infile;
	private File outFile;
	private HSSFWorkbook doc;
	private ArrayList<String> array = new ArrayList<String>();
	private static ArrayList<String> fileNames = new ArrayList<String>();
	private PrintWriter printWriter;
	private Scanner scanner;
	private static File idsFile = null;

	public static void main(String[] args){
		new IDUpdater();
	}

	public IDUpdater(){
		try{
			//Open IDS File
			JFileChooser dialog = new JFileChooser();
			dialog.setDialogTitle("SELECT IDS FILE");
			dialog.setCurrentDirectory(new File("."));
			FileNameExtensionFilter filter = new FileNameExtensionFilter("txt", "txt");
			dialog.setFileFilter(filter);
			int option = dialog.showOpenDialog(new JFrame());
			if(option == JFileChooser.APPROVE_OPTION){
				idsFile = dialog.getSelectedFile();
			}


			scanner = new Scanner(new FileInputStream(idsFile));
			while(scanner.hasNext()){
				array.add(scanner.next());
			}

			//FileChooser
			dialog = new JFileChooser();
			dialog.setDialogTitle("SELECT MATRICES");
			dialog.setCurrentDirectory(new File("."));
			dialog.setMultiSelectionEnabled(true);
			filter = new FileNameExtensionFilter("xls", "xls");
			dialog.setFileFilter(filter);
			option = dialog.showOpenDialog(new JFrame());
			if(option == JFileChooser.APPROVE_OPTION){
				File[] files = dialog.getSelectedFiles();
				fileNames = new ArrayList<String>();
				for(File f : files){
					fileNames.add(f.getAbsolutePath());
				}
			}

			for(String file : fileNames){
				if(file == null){ System.out.println("Nothing to execute"); return; }
				infile = new File(file);
				outFile = new File(file.split(".xls")[0] + "_output.xls");

				in = new FileInputStream(infile);
				out = new FileOutputStream(outFile);
				doc = new HSSFWorkbook(in);

				HSSFSheet sheet = doc.getSheetAt(0);
				int lastRow = sheet.getLastRowNum();

				for(int i = 0; i < lastRow; i++){
					Row row = sheet.getRow(i);
					if(row == null) continue;
					Cell cell8 = row.getCell(8);
					//Cell cell8 = row.getCell(9); // #add to report

					if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;
					String cell8Content = cell8.getStringCellValue();
					if(cell8Content.matches("test case start")){
						Cell cell0 = row.getCell(0);
						if(cell0 == null || cell0.getCellType() != Cell.CELL_TYPE_STRING) continue;
						cell0.setCellValue(array.get(0));
						array.remove(0);
					}
				}

				printWriter = new PrintWriter(new FileOutputStream(idsFile));
				for(int i = 0; i < array.size(); i++){
					printWriter.write(array.get(i) + "\n");
				}

				printWriter.close();
				doc.write(out);
			}
			JOptionPane.showMessageDialog(null, "Done! Find updated matrices in same folder");
		}
		catch(IndexOutOfBoundsException e){ 
			JOptionPane.showMessageDialog(null, "Ids count is not enought to update matrices");
			e.printStackTrace(); 
		}
		catch(FileNotFoundException e){ e.printStackTrace(); }
		catch(IOException e){ e.printStackTrace(); }
		catch(Exception e){ e.printStackTrace(); }
		finally{
			try{
				in.close();
				out.close();
			}
			catch(Exception e){ e.printStackTrace(); }		
			System.exit(0);
		}
	}
}
