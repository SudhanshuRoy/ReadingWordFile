package com.examfactor;

import java.io.FileNotFoundException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ICell;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

public class ReadTable {
	
	// Iterating over image of each cell
	public static void getImage(ICell cell) {
		for (XWPFParagraph p : ((IBody) cell).getParagraphs()) {
			for (XWPFRun run : p.getRuns()) {
				for (XWPFPicture pic : run.getEmbeddedPictures()) {
					byte[] pictureData = pic.getPictureData().getData();

					System.out.println("image : " + pictureData);
				}
			}
		}
	}

	public static void main(String[] args) throws FileNotFoundException {

		String fileName = "sample.docx";

		try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(fileName)))) {

			// whole word document data
			Iterator<IBodyElement> docElementsIterator = doc.getBodyElementsIterator();

			while (docElementsIterator.hasNext()) {

				// All elements of word file table along with others
				IBodyElement docElement = docElementsIterator.next();

				// filter table out of documents
				if ("TABLE".equalsIgnoreCase(docElement.getElementType().name())) {
					// list of all tables
					List<XWPFTable> xwpfTableList = docElement.getBody().getTables();

					// iterating over list of tables
					for (XWPFTable xwpfTable : xwpfTableList) {
						
//						System.out.println("Total Rows : " + xwpfTable.getNumberOfRows());

						System.out.println("#########New Question start##########");
						// iterating over records of one table
						for (int i = 0; i < xwpfTable.getRows().size(); i++) {

							// iterating over cells of one record
							for (int j = 0; j < xwpfTable.getRow(i).getTableCells().size(); j++) {

								System.out.println(xwpfTable.getRow(i).getCell(j).getText());

								if (xwpfTable.getRow(i).getCell(j) != null) {

									getImage(xwpfTable.getRow(i).getCell(j));

								}

							}
							System.out.println("****************");
						}

					}

				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("Done....");

	}


}
