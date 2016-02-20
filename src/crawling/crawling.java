package crawling;
import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

public class crawling extends JFrame implements ActionListener{
	static Calendar cal = Calendar.getInstance();
	static int year = cal.get ( cal.YEAR );

	static String month;
	static String date;
	static String today ;
	JPanel background;
	JPanel[] p;
	static JTextField num , num1;
	JButton ok, exit;
	JLabel month_label, date_label;
	ImageIcon icon;

	public crawling(){ 
		setSize(370,200);
		Dimension frameSize = getSize();
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		setLocation((screenSize.width - frameSize.width)/2,
				(screenSize.height - frameSize.height)/2);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setTitle("동업계 방송편성표 크롤링 프로그램");
		setResizable(false);
		setLayout(null);
		setVisible(true);

		icon = new ImageIcon("img/hmall.png");
		background = new JPanel() {
			public void paintComponent(Graphics g) {
				g.drawImage(icon.getImage(), 0, 0, null);
				setOpaque(false);
				super.paintComponent(g);
			}
		};
		add(background);
		background.setBounds(0, 0, 400, 300);

		background.setLayout(null);

		month_label = new JLabel("월(月)을 입력하세요.");
		background.add(month_label);
		month_label.setBounds(50,20,150,25);

		num = new JTextField();
		background.add(num);
		num.setBounds(250,20,50,25);
		num.requestFocus();

		date_label = new JLabel("일(日)을 입력하세요.");
		background.add(date_label);
		date_label.setBounds(50,50,150,25);

		num1 = new JTextField();
		background.add(num1);
		num1.setBounds(250,50,50,25);
		num1.requestFocus();

		ok = new JButton("생성(O)");
		background.add(ok);
		ok.setBounds(70,90,100,25);

		exit = new JButton("나가기(Q)");
		background.add(exit);
		exit.setBounds(190,90,100,25);


		ok.addActionListener(this);
		exit.addActionListener(this);

		ok.setMnemonic('O');
		exit.setMnemonic('Q');
	}

	public void actionPerformed(ActionEvent e) {
		if (e.getSource() == ok) {
			if (getMonth().trim().length() != 0 && getDate().trim().length() != 0) {
				start();
				num.requestFocus();
				num.setText("");
				num1.requestFocus();
				num1.setText("");
				JOptionPane.showMessageDialog(this,"'"+year+"년"+month+"월"+date+"일_"+"방송편성표.xls' 파일이 생성되었습니다.","파일 생성",JOptionPane.INFORMATION_MESSAGE);
			} else {
				JOptionPane.showMessageDialog(null, "빈칸을 입력해 주세요.");
			}
		} else if (e.getSource() == exit) {
			System.exit(0);
		}
	}

	public static String getMonth() {
		return num.getText();
	}

	public static String getDate() {
		return num1.getText();
	}

	public static void main(String[]args){

		try {
			UIManager.setLookAndFeel("javax.swing.plaf.nimbus.NimbusLookAndFeel");
			new crawling();
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (InstantiationException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (UnsupportedLookAndFeelException e) {
			e.printStackTrace();
		}
	}

	private static void start(){
		month = crawling.getMonth();
		date = crawling.getDate();
		today = String.valueOf(year) + ( Integer.parseInt(month) < 10 ? "0"+month : month ) +( Integer.parseInt(date) < 10 ? "0"+date : date );

		Workbook wb = new HSSFWorkbook();

		Font font = wb.createFont();
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);

		Font font1 = wb.createFont();
		font1.setBoldweight(Font.BOLDWEIGHT_BOLD);
		font1.setFontHeightInPoints((short)14);

		CellStyle style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setFont(font);

		CellStyle style2 = wb.createCellStyle();
		style2.setBorderBottom(CellStyle.BORDER_THIN);
		style2.setBorderLeft(CellStyle.BORDER_THIN);
		style2.setBorderRight(CellStyle.BORDER_THIN);
		style2.setBorderTop(CellStyle.BORDER_THIN);
		style2.setAlignment(CellStyle.ALIGN_CENTER);
		style2.setFont(font1);

		try{
			// GS SHOP
			Sheet sheet = wb.createSheet("GSshop");
			sheet.setColumnWidth(0, 2*256);
			sheet.setColumnWidth(1, 15*256);
			sheet.setColumnWidth(2, 13*256);
			sheet.setColumnWidth(3, 13*256);
			sheet.setColumnWidth(4, 20*256);
			sheet.setColumnWidth(5, 110*256);

			HashMap<String, ArrayList<String>> mall = gs();
			writeExcel(mall,sheet,style,style2,"GSshop");

			// CJ 오쇼핑
			Sheet sheet2 = wb.createSheet("CJ오쇼핑");
			sheet2.setColumnWidth(0, 2*256);
			sheet2.setColumnWidth(1, 15*256);
			sheet2.setColumnWidth(2, 13*256);
			sheet2.setColumnWidth(3, 22*256);
			sheet2.setColumnWidth(4, 15*256);
			sheet2.setColumnWidth(5, 38*256);
			HashMap<String, ArrayList<String>> cjmall = cj();
			writeExcel(cjmall,sheet2,style,style2,"CJ오쇼핑");

			// 롯데 홈쇼핑
			Sheet sheet3 = wb.createSheet("롯데홈쇼핑");
			sheet3.setColumnWidth(0, 2*256);
			sheet3.setColumnWidth(1, 15*256);
			sheet3.setColumnWidth(2, 13*256);
			sheet3.setColumnWidth(3, 13*256);
			sheet3.setColumnWidth(4, 15*256);
			sheet3.setColumnWidth(5, 110*256);
			HashMap<String, ArrayList<String>> lottemall = lotte();
			writeExcel(lottemall,sheet3,style,style2,"롯데홈쇼핑");


			FileOutputStream fileOut = new FileOutputStream("./"+year+"년"+month+"월"+date+"일_"+"방송편성표.xls");
			wb.write(fileOut);
			fileOut.close();
			System.out.println("'"+year+"년"+month+"월"+date+"일_"+"방송편성표.xls' 파일이 생성되었습니다.");
		}catch(Exception e){
			e.printStackTrace();
		}
	}

	public static void writeExcel(HashMap<String, ArrayList<String>> mall , Sheet sheet, CellStyle style,CellStyle style2,String mallName){
		Row title = sheet.createRow(1);

		HSSFCell cell = (HSSFCell) title.createCell((short)1);
		cell.setCellStyle(style2);
		cell.setCellValue(mallName);

		Row name = sheet.createRow(2);

		cell = (HSSFCell) name.createCell((short)1);
		cell.setCellStyle(style2);
		cell.setCellValue(year+"/"+month+"/"+date);

		cell = (HSSFCell) name.createCell((short)2);
		cell.setCellStyle(style);
		cell.setCellValue("시간");

		cell = (HSSFCell) name.createCell((short)3);
		cell.setCellStyle(style);
		cell.setCellValue("카테고리");

		cell = (HSSFCell) name.createCell((short)4);
		cell.setCellStyle(style);
		cell.setCellValue("가격");

		cell = (HSSFCell) name.createCell((short)5);
		cell.setCellStyle(style);
		cell.setCellValue("상품");

		for(int i = 0 ; i < mall.get("time").size() ; i++ ){
			Row row = sheet.createRow((short)3+i);
			for(int j = 2 ; j < 3 ; j++){
				row.createCell(j).setCellValue(mall.get("time").get(i));
				if(mallName.equals("CJ오쇼핑") || mallName.equals("GSshop")){
					row.createCell(j+1).setCellValue(mall.get("category").get(i));
				}
				row.createCell(j+2).setCellValue(mall.get("price").get(i));
				row.createCell(j+3).setCellValue(mall.get("content").get(i));
			}
		}
	}

	// GS SHOP
	public static HashMap <String, ArrayList<String>> gs() throws IOException{
		String time = null;
		String category = null;
		String price = null;
		String content = null;

		Elements temp = null;

		HashMap<String, ArrayList<String>> gs = new HashMap<String, ArrayList<String>>();
		ArrayList<String> gsTime = new ArrayList<String>();
		ArrayList<String> gsCategory = new ArrayList<String>();
		ArrayList<String> gsContent = new ArrayList<String>();
		ArrayList<String> gsPrice = new ArrayList<String>();

		String address="http://with.gsshop.com/tv/tvScheduleMain.gs?selectDate=" +today;

		Document doc = Jsoup.connect(address).get();
		temp = doc.getElementsByClass("time");

		for(int i = 0 ; i < temp.size() ; i++){
			time = temp.get(i).getElementsByClass("tdWrap").text().replace("-", " ~ ");
			category = temp.get(i).parent().getElementsByClass("desc").get(0).getElementsByClass("category").text();
			content = temp.get(i).parent().getElementsByClass("desc").get(0).getElementsByTag("a").text().replace(category, "");
			price = temp.get(i).parent().getElementsByClass("price").get(0).getElementsByTag("b").text().replace(" "," / ") + "원";
			gsTime.add(time);
			gsCategory.add(category);
			gsContent.add(content);
			gsPrice.add(price);
		}
		gs.put("time", gsTime);
		gs.put("category", gsCategory);
		gs.put("price", gsPrice);
		gs.put("content", gsContent);

		return gs;
	}

	// CJ 오쇼핑
	public static HashMap <String, ArrayList<String>> cj() throws IOException{
		String time = null;
		String category = null;
		String price = null;
		String content = null;

		Elements temp = null;

		HashMap<String, ArrayList<String>> cj = new HashMap<String, ArrayList<String>>();
		ArrayList<String> cjTime = new ArrayList<String>();
		ArrayList<String> cjCategory = new ArrayList<String>();
		ArrayList<String> cjPrice = new ArrayList<String>();
		ArrayList<String> cjContent = new ArrayList<String>();

		String address = "http://www.cjmall.com/etv/broad/broad_list_ajax.jsp?insert_date="+today;

		Document doc = Jsoup.connect(address).get();
		temp = doc.select(".e_item");
		for(int i = 0 ; i < temp.size() ; i++){
			if(temp.get(i).select(".viewGuest").hasText()){
				time = temp.get(i).select(".onair_info").get(0).select("strong").text();
			}else{
				time = "";
			}
			category = temp.get(i).select(".onair_info").get(0).getElementsByTag("b").text().replace("...", "");
			price = temp.get(i).getElementsByClass("e_cost").text();
			content = temp.get(i).getElementsByClass("e_copy").text();
			if(i != 0 && time.equals(temp.get(i-1).select(".onair_info").get(0).getElementsByTag("strong").text())){
				cjTime.add("");
			}else{
				cjTime.add(time);
			}
			cjPrice.add(price);
			cjCategory.add(category);
			cjContent.add(content);
		}
		cj.put("time", cjTime);
		cj.put("price", cjPrice);
		cj.put("category", cjCategory);
		cj.put("content", cjContent);

		return cj;
	}

	// 롯데 홈쇼핑
	public static HashMap <String, ArrayList<String>> lotte() throws IOException{
		String time = null;
		String category = null;
		String price = null;
		String content = null;

		Elements temp = null;

		HashMap<String, ArrayList<String>> lotte = new HashMap<String, ArrayList<String>>();
		ArrayList<String> lotteTime = new ArrayList<String>();
		ArrayList<String> lotteCategory = new ArrayList<String>();
		ArrayList<String> lottePrice = new ArrayList<String>();
		ArrayList<String> lotteContent = new ArrayList<String>();

		String address = "http://www.lotteimall.com/main/searchTvPgmByDay.lotte?bd_date="+today;
		Document doc = Jsoup.connect(address).get();
		temp = doc.getElementsByClass("rn_tsitem_wrap");

		for(int i = 0 ; i < temp.size() ; i++){
			time = temp.get(i).getElementsByClass("rn_tsitem_time").text();
			content = temp.get(i).getElementsByClass("rn_tsitem_info").get(0).getElementsByTag("a").text();
			price = temp.get(i).getElementsByClass("rn_tsitem_priceinfo").text();
			lotteTime.add(time);
			lottePrice.add(price);
			//lotteCategory.add(category);
			lotteContent.add(content);
		}
		lotte.put("time", lotteTime);
		lotte.put("price",lottePrice);
		lotte.put("content", lotteContent);
		return lotte;
	}
}