package org.red1.ninja2;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.sql.Types;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import org.adempiere.exceptions.AdempiereException;
import org.adempiere.exceptions.DBException;
import org.adempiere.util.ServerContext;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.compiere.Adempiere;
import org.compiere.model.MColumn;
import org.compiere.model.MField;
import org.compiere.model.MMenu;
import org.compiere.model.MRefList;
import org.compiere.model.MReference;
import org.compiere.model.MTab;
import org.compiere.model.MTable;
import org.compiere.model.MTree_Base;
import org.compiere.model.MTree_NodeMM;
import org.compiere.model.MWindow;
import org.compiere.model.M_Element;
import org.compiere.model.PO;
import org.compiere.model.Query;
import org.compiere.model.X_AD_Menu;
import org.compiere.model.X_AD_Reference;
import org.compiere.model.X_AD_TreeNodeMM;
import org.compiere.util.AdempiereUserError;
import org.compiere.util.CLogger;
import org.compiere.util.DB;
import org.compiere.util.DisplayType;
import org.compiere.util.Env;
import org.compiere.util.Ini;
import org.compiere.util.Trx;
import org.compiere.util.ValueNamePair;
import org.eclipse.equinox.app.IApplication;
import org.eclipse.equinox.app.IApplicationContext;

public class ModuleMaker implements IApplication {
	private String File_Directory = "/home/red1/hongkong/Ninja.xls";
	int p_ValueLength = 22;
	int p_NameLength = 22;
	HSSFWorkbook workbook = null;  
	int reference = 10; //String
	List<Integer> refset = new ArrayList<Integer>();
	String trxName = "";
	int ref_cnt = 0;
	int model_cnt = 0;
	int listvaluesadded=0;
	String prefix = "CA_";//shall fetch first two letters to UPPER case of Menu Title in first Excel cell.
	String entity = "A";
	String ad_name;
	List<String> masterset = new ArrayList<String>();
	List<MTable> tableset = new ArrayList<MTable>();
	List<MWindow> windowset = new ArrayList<MWindow>();
	List<MTab> tabset = new ArrayList<MTab>();
	boolean synchronize = false;
	MMenu mainMenu;
	
	@Override
	public Object start(IApplicationContext context) throws Exception {
		/** Initialise context for the current thread*/
        
		startADempiere();
        
		for (int i=0;i<workbook.getNumberOfSheets();i++) {
			HSSFSheet sheet = workbook.getSheetAt(i);
			String sheetName = sheet.getSheetName();

			if ("List".equals(sheetName)) 
				doList(sheet);
			else if ("Model".equals(sheetName)) 
				doModel(sheet);  
			else if (sheetName.split("_").length>1)
				importModel(sheet);
 		} 
		
		endADempiere();
		
		return null;
	}

	private void importModel(HSSFSheet sheet) {
		// TODO Auto-generated method stub
		HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
		MTable table = new Query(Env.getCtx(), MTable.Table_Name, MTable.COLUMNNAME_TableName+"=?",trxName)
				.setParameters(sheet.getSheetName())
				.first();
		if(table==null)
			throw new AdempiereException("**TABLE DOES NOT EXIST, DEFINE FIRST IN MODEL SHEET : "+sheet.getSheetName());
		//count number of records already inserted
		String count = "SELECT COUNT(*) FROM "+table.getTableName(); 
		int no = DB.getSQLValue(trxName, count);
		int PK = 1000000+no;
		
		count = "SELECT AD_Client_ID FROM AD_Client order by AD_Client_ID DESC LIMIT 1";
		int AD_Client_ID = DB.getSQLValue(trxName, count);
		
		StringBuilder colset = new StringBuilder("INSERT INTO "+table.getTableName()
		+" (AD_Client_ID,AD_Org_ID,IsActive,Created,CreatedBy,Updated,UpdatedBy,"+table.getTableName()+"_ID," );
		StringBuilder dataset = new StringBuilder(" VALUES ("+AD_Client_ID+","+"0,'Y','"
		+table.getUpdated()+"',"+table.getUpdatedBy()+",'"+table.getUpdated()+"',"+table.getUpdatedBy()+",");
		StringBuilder colsetadd = new StringBuilder();
		StringBuilder datasetadd = new StringBuilder();
		List<MColumn>columns = new ArrayList<MColumn>();
		
		for (int i=0;i<=sheet.getLastRowNum();i++) {
			HSSFRow row = sheet.getRow(i);
			if (row==null)
				break;
			if (i>0){
				datasetadd.append(PK+",");
				PK++;
			}
			for (int y=0;y<row.getLastCellNum();y++) {
				HSSFCell cell = row.getCell(y);
				if (cell==null)
					break;
				if (i==0) { //columns row one
					MColumn column = new Query(Env.getCtx(),MColumn.Table_Name,MColumn.COLUMNNAME_Name+"=? AND "+MColumn.COLUMNNAME_AD_Table_ID+"=?",trxName)
							.setParameters(cell.getStringCellValue(),table.get_ID())
							.first();
					if (column==null) {
						throw new AdempiereException("COLUMN NAME >"+cell.getStringCellValue()+"< NOT FOUND IN TABLE > "+sheet.getSheetName());
					}
					columns.add(column);
					colsetadd.append(column.getColumnName()+",");
				}else { //data rows
					String cellvalue = dataFormatter.formatCellValue(cell);  
					 if(cell.getCellType() == CellType.FORMULA) {
			            	switch(cell.getCachedFormulaResultType()) {
		                    	case NUMERIC:
		                    		cellvalue= Double.toString(cell.getNumericCellValue());
		                    		break;
		                    	case STRING:
		                    		cellvalue= cell.getRichStringCellValue().getString();
		                    		break;
							default:
								break;
		                }
		            } 	 
					int reference = columns.get(y).getAD_Reference_ID();
					if (reference==10)
						datasetadd.append("'"+cellvalue+"',");
					else
						datasetadd.append(cellvalue+",");
				}
			}
			if (i==0) {
				colsetadd.deleteCharAt(colsetadd.length()-1).append(")");
				continue;
			}
			datasetadd.deleteCharAt(datasetadd.length()-1).append(")");
			String SQL = colset.toString()+colsetadd.toString()+dataset.toString()+datasetadd.toString();
			DB.executeUpdateEx(SQL,trxName);
			System.out.println(SQL);
			datasetadd = new  StringBuilder();
		}
	}

	private void startADempiere() throws Exception{
		Properties serverContext = new Properties();
	    ServerContext.setCurrentInstance(serverContext);
	
	    String propertyFile = Ini.getFileName(false);
	    File prop = new File(propertyFile);
	    if (!prop.exists()) {
	    	throw new IllegalStateException("idempiere.properties file missing. Path="+prop.getAbsolutePath());
	    }
	    if (!Adempiere.isStarted())
	    {
	        boolean started = Adempiere.startup(false);
	        if(!started)
	        {
	            throw new AdempiereException("Could not start ADempiere");
	        }
	    }
	
		//check data
	    ref_cnt = new Query(Env.getCtx(), X_AD_Reference.Table_Name, "", null).count();
	    model_cnt = new Query(Env.getCtx(),MTable.Table_Name, MTable.COLUMNNAME_EntityType+"=?", null)
	    		.setParameters("A").count();
		trxName = Trx.createTrxName("Ninja"); 
		
		FileInputStream file = new FileInputStream(File_Directory);  
		workbook = new HSSFWorkbook(file); 
		Runtime.getRuntime().exec("clear");   
		System.out.println(+workbook.getNumberOfSheets()+" < worksheets in "+File_Directory);
		System.out.println(ref_cnt+" < Reference Table Count --- SYSTEM LAUNCHED --- TRANSACTION NAME = "+trxName);
		System.out.println(model_cnt+" < App Entity");
	}

	private void endADempiere() throws IOException {
		workbook.close();
		
		Trx trx = Trx.get(trxName, false);
		trx.commit();
		trx.close();
		
	}

	@Override
	public void stop() { 
		
	}
	
	private void doList(HSSFSheet sheet) { 
		int rowcount = sheet.getLastRowNum()+1;
		for (int i=0;i<rowcount;i++) {
			HSSFRow row = sheet.getRow(i);
			
			int colcount = row.getLastCellNum();
			
			if (i==0) {
				for (int y=0;y<colcount;y++) {
					refset.add(createReference(row.getCell(y)));
				}
			}
			else
				for (int y=0;y<colcount;y++) {
					createRefList(row.getCell(y),y);
				}
		}
		ref_cnt = new Query(Env.getCtx(), X_AD_Reference.Table_Name, "", trxName).count();
		System.out.println(ref_cnt+" < Reference Table Count --- SYSTEM COMPLETED --- TRANSACTION NAME = "+trxName);
		System.out.println(listvaluesadded+" < ReferenceList Table values added");
	}

	private void doModel(HSSFSheet sheet) throws Exception {  
		HSSFRow firstrow = sheet.getRow(0);
		HSSFCell firstcell = firstrow.getCell(0);
		if (firstcell==null)
			throw new AdempiereException("Put Main Menu name in A1 - first cell in Ninja Excel. First two chars is PREFIX i.e. HR_");
		mainMenu = createMainMenu(firstcell); 
		
		int rowcount = sheet.getLastRowNum()+1;
		for (int i=0;i<rowcount;i++) {
			HSSFRow row = sheet.getRow(i);
			
			int colcount = row.getLastCellNum();
			{ 
				for (int y=0;y<colcount;y++) { 
					if (row.getCell(y)==null) {
						masterset.add("");
						continue;
					}
						
					String cellvalue = row.getCell(y).getStringCellValue();	
					if (cellvalue.split("_").length>1)
						ad_name=cellvalue;
					else
						ad_name = prefix+(cellvalue).trim().replaceAll("\\s+","");	
					if (i==0) 
						if (y==0)
							masterset.add("");
						else
							masterset.add(cellvalue);
					else if (i==1) {
						if (y==0 || masterset.get(y).isBlank())
							windowset.add(createWindow(cellvalue));
						else {
							//set windowset to masterset value
							MWindow window = new Query(Env.getCtx(),MWindow.Table_Name,MWindow.COLUMNNAME_Name+"=? AND "
									+MWindow.COLUMNNAME_AD_Window_ID+">?",trxName)
									.setParameters(masterset.get(y),999999).first();
							windowset.add(window);
						}
						createTable(cellvalue,y); //tableset.added
						tabset.add(createTab(cellvalue,y)); 
						model_cnt++;
					}
					else {
						if (!cellvalue.isBlank())
							addColumn(y, cellvalue); 
					}
				}
			}
		}
		 for (int i=0;i<tabset.size();i++) {
			 createTabFields(tabset.get(i));
		 }
		 System.out.println(Synchronize()+" < App Entity Added");
	}

	private MMenu createMainMenu(HSSFCell firstcell) { 
		String menuname = firstcell.getStringCellValue();
		prefix = menuname.substring(0,2).toUpperCase()+"_";
		
		MMenu menu = new Query(Env.getCtx(),MMenu.Table_Name,MMenu.COLUMNNAME_Name+"=?",trxName)
				.setParameters(menuname)
				.first();
		if (menu==null) {
			menu = new MMenu(Env.getCtx(), 0, trxName);
			menu.setName(menuname);
			menu.setAction(X_AD_Menu.ACTION_Window);
			menu.setEntityType(entity);
			menu.set_ValueOfColumn(MMenu.COLUMNNAME_AD_Client_ID, 0);
			menu.setAD_Org_ID(0);
			menu.setIsSummary(true);
			menu.saveEx(trxName); 
		}				
		return menu; 
	}

	private MMenu createMenu(MWindow window) { 
		MMenu menu = new Query(Env.getCtx(),MMenu.Table_Name,MMenu.COLUMNNAME_AD_Window_ID+"=?",trxName)
				.setParameters(window.getAD_Window_ID())
				.first();
		if (menu==null) {
			menu = new MMenu(Env.getCtx(), 0, trxName);
			menu.setName(window.getName());
			menu.setAD_Window_ID(window.getAD_Window_ID());
			menu.setAction(X_AD_Menu.ACTION_Window);
			menu.setEntityType(entity);
			menu.set_ValueOfColumn(MMenu.COLUMNNAME_AD_Client_ID, 0);
			menu.setAD_Org_ID(0);
			menu.saveEx(trxName); 
		}				
		return menu; 
	}

	private MTable createTable(String cellvalue, int y) throws Exception {
		MTable table = new Query(Env.getCtx(), MTable.Table_Name,MTable.COLUMNNAME_TableName+"=? AND "
				+MTable.COLUMNNAME_AD_Table_ID+">?",trxName)
				.setParameters(ad_name,999999)
				.first();
		if (table!=null) {
			tableset.add(table);
			return table;
		}
		cellvalue = cellvalue.replaceAll("\\s+","").trim();
		cellvalue = leadingCapsSpacing(cellvalue);
		
		M_Element id = M_Element.get (Env.getCtx(),ad_name + "_ID",trxName);
		if (id==null) {
			id = new M_Element(Env.getCtx(), 0, trxName);
			id.setColumnName(ad_name + "_ID");
			id.setPrintName(cellvalue);
			id.setName(cellvalue);
			//id.setDescription(mt.getDescription());
			//id.setHelp(mt.getHelp());
			id.setEntityType(entity);
		
			try {
				id.saveEx(trxName);
			} catch (Throwable t) {
				System.out.println("AD_Element "+ad_name + "_ID" + " already exists ?");
			}
		}
		//UUID column
		M_Element elementUU = M_Element.get (Env.getCtx(), ad_name + "_UU");
		if (elementUU == null){
			elementUU = new M_Element (Env.getCtx(), ad_name + "_UU", "U", trxName);
			elementUU.saveEx(trxName);
		}
		table = new MTable(Env.getCtx(), 0, trxName);
		table.setTableName(ad_name);
		table.setName(cellvalue);
		table.setDescription("");
		table.setEntityType(entity);
		table.setAccessLevel("3");
		table.set_ValueOfColumn("AD_Client_ID", 0);
		table.saveEx(trxName);
		tableset.add(table);
		//ID Column
		MColumn columnID = new MColumn(Env.getCtx(), 0, trxName);
		columnID.setColumnName(ad_name+"_ID");
		columnID.setName(cellvalue);
		columnID.setAD_Element_ID(id.get_ID());
		columnID.setAD_Reference_ID(13); //"ID"
		columnID.setFieldLength(22);
		columnID.setIsMandatory(true);
		columnID.setIsKey(true);
		columnID.setAD_Table_ID(table.get_ID());
		columnID.setEntityType(entity);
		columnID.saveEx(trxName);
	
		String copylist = "'AD_Client_ID','AD_Org_ID','Created','CreatedBy','UpdatedBy','Updated','IsActive'";
	
		createColumns(copylist,y);
	
		return table; 
	}

	private MWindow createWindow(String cellvalue) {  
		MWindow window = new Query(Env.getCtx(),MWindow.Table_Name,MWindow.COLUMNNAME_Name+"=?",trxName)
				.setParameters(cellvalue).first();
		if (window==null) {
			window = new MWindow(Env.getCtx(),0,trxName);
			window.setName(leadingCapsSpacing(cellvalue)); 
			window.setWindowType("M");
			window.setEntityType(entity);
			window.saveEx(trxName);

			attachMenuTree(createMenu(window));
		}
		return window;
	}

	private void attachMenuTree(MMenu menu) { 
		MTree_Base menutree =  new MTree_Base(Env.getCtx(),10,trxName);  
		X_AD_TreeNodeMM mm = new Query(Env.getCtx(),MTree_NodeMM.Table_Name,MTree_NodeMM.COLUMNNAME_AD_Tree_ID+"=? AND "+MTree_NodeMM.COLUMNNAME_Node_ID+"=?",trxName)
		.setParameters(menutree.get_ID(),menu.get_ID())
		.first();
		if (mm != null){
			mm.setParent_ID(mainMenu.get_ID());
			mm.setAD_Org_ID(0);
			mm.saveEx(trxName);
		}
	}

	private MTab createTab(String cellvalue, int y) throws Exception {
		MTable table = tableset.get(y);
		int holder=y;
		MTab tab = new Query(Env.getCtx(),MTab.Table_Name,MTab.COLUMNNAME_AD_Table_ID+"=? AND "
				+MTab.COLUMNNAME_AD_Tab_ID+">?",trxName)
				.setParameters(table.get_ID(),999999)
				.first();
				if (tab==null){
					if (windowset.get(y)==null) {
						//seems to be grand child, retreat one column
						holder--;
						tab = tabset.get(holder);
						if (!tab.getName().equals(tableset.get(holder).getName()))
							throw new AdempiereException(cellvalue+" ** CHECK YOUR EXCEL TABLE IS PAPA,ME,CHILD ARRANGEMENT");
						tab = new MTab(windowset.get(holder));
					}else
						tab = new MTab(windowset.get(y));
					tab.setName(leadingCapsSpacing(cellvalue));
					tab.setAD_Table_ID(table.getAD_Table_ID());
 					int wintabs = new Query(Env.getCtx(),MTab.Table_Name,MTab.COLUMNNAME_AD_Window_ID+"=?",trxName)
					.setParameters(windowset.get(holder).get_ID())
					.count();
					tab.setSeqNo(wintabs*10+10);
					tab.setEntityType(entity);
					int tablevel = 0;
					if (masterset.get(y).isBlank())
						;
					else {
						MTab tablast = new Query(Env.getCtx(), MTab.Table_Name,MTab.COLUMNNAME_Name+"=?",trxName)
								.setParameters(masterset.get(y)).first();
						tablevel = tablast.getTabLevel()+1;
					}
					tab.setTabLevel(tablevel);
					//
					if (wintabs>0) {
						String parentname = masterset.get(y);
						String parentlink = ""; 
						if (cellvalue.contains(">")) {
							String[] split = cellvalue.split(">");
							parentlink = split[1];
						}
						int parentLinkID = setForeignKey(parentname, parentlink,
								tableset.get(y),y); 
						if (parentLinkID>0){
							tab.setAD_Column_ID(parentLinkID);
							tab.setParent_Column_ID(parentLinkID);
						}
					}
					tab.saveEx(trxName);
				}
		return tab;
	}

	/**
	 * create columns and synchronize
	 * @param enttype
	 * @param trxName
	 * @param copylist
	 * @return
	 * @throws Exception
	 */
	private MTable createColumns(String copylist, int y)
			throws Exception {
		// 102 - AD_Reference
		String sql = "SELECT AD_COLUMN_ID from AD_COLUMN where ad_table_ID=102 and columnname IN (" + copylist + ")";
		PreparedStatement ps = null;
		ResultSet rs = null;
		try
		{
			ps = DB.prepareStatement(sql.toString(), trxName);
			rs = ps.executeQuery();

			while (rs.next())
			{
				MColumn col = new MColumn(Env.getCtx(), rs.getInt(1), trxName);
				MColumn dest = new MColumn(Env.getCtx(), 0, trxName);
				PO.copyValues(col, dest);
				dest.set_ValueNoCheck("AD_Table_ID",  tableset.get(y).get_ID());
				dest.set_ValueNoCheck("EntityType", entity);
				dest.setIsMandatory(false);
				dest.setDefaultValue("");
				try {
					dest.saveEx(trxName);
				} catch (Throwable t) {
					System.out.println("Duplicate Column ? "+dest.getColumnName() + " - " +tableset.get(y).getTableName());
				}
				dest.setIsMandatory(col.isMandatory());
				dest.setDefaultValue(col.getDefaultValue());
				dest.setDescription("");
				dest.setHelp("");
				try {
					dest.saveEx(trxName);
				} catch (Throwable t) {
					System.out.println("Duplicate Column ? "+dest.getColumnName() + " - " +tableset.get(y).getTableName());
				}
			}
		}
		catch (SQLException e)
		{
			throw new DBException(e, sql);
		}
		finally
		{
			DB.close(rs, ps);
			rs = null; ps = null;
		}
		return tableset.get(y);
	}
	 
	/**
	 * Default or S# String datatype with length = 22
	 * if ElementName exists, its datatype adopted
	 * Q# = Qty - 11
	 * Y# = Yes/No - 1
	 * A# = Amount - 11
	 * D# = DateTime - 11
	 * T# = Text - 2000 
	 * L# - List Validation /Table/Data
	 * @param colname
	 * @param trxName
	 */
	private MColumn addColumn(int y , String cellvalue){ 
		String name = cellvalue;
		String[] split = cellvalue.split("#");
		String type="";
		if (split.length==2){
			type = split[0].trim();
			name = split[1];
		}
		String colname = name.replaceAll("\\s+","");
		MTable table = tableset.get(y);
		String[] changename = colname.split(">");
		if (changename.length>1){
			colname = changename[0];
		}
		
		int reference = 10; //String
		int length = 22;
		if (colname.endsWith("_ID"))
			if (!colname.contains(prefix)&&(colname.split("_").length<3))
				colname = prefix+colname;
		MColumn column = table.getColumn(colname);
		if (column==null)
			column = new MColumn(Env.getCtx(),0,trxName);
		if (colname.isBlank())
			System.out.println(" -- COLNAME IS BLANK -- ");
		//element existed?
		M_Element element = M_Element.get(Env.getCtx(), colname, trxName);
		//new element
		if (element==null){
			element = new M_Element (Env.getCtx(), colname, entity, trxName);
			element.saveEx(trxName);
		} 
		column.setAD_Element_ID(element.get_ID());
		column.setAD_Reference_ID(reference); 
		column.setColumnName(colname);
		column.setAD_Table_ID(table.getAD_Table_ID());
		column.setEntityType(entity); 
		column.setName(leadingCapsSpacing(name)); 
		column.setFieldLength(length);
		column.setIsMandatory(false);
		column.setIsKey(false);	
	
		if (type.equals("Q")||column.toString().contains("Qty")) {//Quantity
			reference = 29;
			length = 11;
		}
		else if (type.equals("Y")||column.toString().startsWith("Is")) {//Boolean Yes/No
			reference = 20;		
			length = 1;
		}			
		else if (type.equals("A")||column.toString().contains("Amt")) {//Amount
			reference = 12;	
			length = 11;
		}
		else if (type.equals("D")||column.toString().contains("Date")) //Date
			reference = 15;	
		else if (type.equals("d")||column.toString().contains("Time")) //Date
			reference = 16;
		else if (type.equals("T")||column.toString().endsWith("Text")||column.toString().startsWith("Text")) {//Text - 2000 chars
			reference = 14;	 
			length = 2000;
		}else if (type.equals("L")) {
			reference = 17;
			MReference refcheck = new Query(Env.getCtx(),X_AD_Reference.Table_Name,X_AD_Reference.COLUMNNAME_Name+"=?", trxName)
			.setParameters(colname)
			.first();
			if (refcheck==null)
				System.out.println(colname+"--REFERENCE Does Not Exist. Please see List in Excel.");
			else
				column.setAD_Reference_Value_ID(refcheck.get_ID()); 
		}
		column.setFieldLength(length); 
		column.setAD_Reference_ID(reference); column.getName();
		if (column.getColumnName().endsWith("_ID") && reference==10){ 
			column.setAD_Reference_ID(DisplayType.TableDir);
			//remove suffix _ID
			String nosuffix = column.getName();
			column.setName(nosuffix.substring(0, nosuffix.length()-3));
		}
		if (type.isEmpty() || type.equals(""))
			column = CalloutProcessor(column);//if element is common, retrieve its settings 
		column.saveEx(trxName);
		
		if (changename.length>1)
			column.setDescription("**"+changename[1]);//pass to during window field creation to swap name
		
		column.setIsUpdateable(true);
		column.saveEx(trxName); //double save to avoid new status over-write at MColumn.java line 332 Sync Terminology
		
		return column;
	}
	
	void createTabFields(MTab tab) throws SQLException {
		int ten = 10;	
		boolean toggle=false; 
			String sql = "SELECT * FROM AD_Column c "
					+ "WHERE AD_Table_ID=?"	
					+ " AND NOT (Name LIKE 'Created%' OR Name LIKE 'Updated%')"
					+ " AND IsActive='Y' "
					+ "ORDER BY AD_Column_ID";
			PreparedStatement pstmt = null;
			ResultSet rs = null;

			pstmt = DB.prepareStatement (sql, trxName);
			pstmt.setInt (1, tab.getAD_Table_ID());
			rs = pstmt.executeQuery ();
			int seqNo = 0;
			while (rs.next ())
			{
				MColumn column = new MColumn (Env.getCtx(), rs, trxName); 
				String columnName = column.getColumnName();
				//
				MField field = new Query(Env.getCtx(),MField.Table_Name,MField.COLUMNNAME_AD_Column_ID+"=? AND "
						+MField.COLUMNNAME_AD_Tab_ID+"=?",trxName)
						.setParameters(column.get_ID(),tab.get_ID())
						.first();
				if (field==null) {
					field = new MField (tab); 
					field.setColumn(column);
					if (column.isKey()){
						field.setIsDisplayed(false);
						field.setIsDisplayedGrid(false);
					}
					// S�quence & IsSameLine
					if (columnName.equals("AD_Client_ID")){
						seqNo = 3;
						field.setIsDisplayedGrid(false);
					}
						
					else if (columnName.equals("AD_Org_ID"))
					{
						seqNo = 6;
						field.setIsSameLine(true);
						field.setIsDisplayedGrid(false);
						field.setXPosition(3);
					}
					else {
						seqNo = seqNo+ten;	
						//toggling X Position
						if (toggle)
							field.setXPosition(3);
						if (toggle)
							toggle=false;
						else 
							toggle=true;
											
					}
					if (columnName.equals("IsActive") && p_ValueLength > 0) {
						field.setIsSameLine(true);	 
						field.setIsDisplayedGrid(false);
					}	
					
					
					String fieldname =  column.getName();
					
					String changename = column.getDescription();
					if (changename!=null && changename.startsWith("**")){
						 fieldname =  changename.substring(2); 
					}
					field.setIsCentrallyMaintained(false);
					field.setName(leadingCapsSpacing(fieldname));	
					field.setSeqNo(seqNo);
					if (field.getAD_Client_ID()!=0)
						field.set_ValueOfColumn("AD_Client_ID", 0);
					field.saveEx(trxName); 
				}
			} 
	}
	
	int Synchronize() throws Exception
	{ //TODO if table no change, then remove table for no synch to happen
		int tablesSynched = 0;
		for (int y=0;y<tableset.size();y++) {
			MTable table = tableset.get(y);
			if (table == null)
				continue;
			tablesSynched++;
			StringBuilder allSql = new StringBuilder();
			List<MColumn> columns = new Query(Env.getCtx(), MColumn.Table_Name, MColumn.COLUMNNAME_AD_Table_ID+"=?", trxName)
					.setParameters(table.get_ID())
					.list();

		  for (MColumn column : columns) {
			//	Find Column in Database
			Connection conn = null;
			ResultSet rs = null;
			try {
				conn = DB.getConnection();
				DatabaseMetaData md = conn.getMetaData();
				String catalog = DB.getDatabase().getCatalog();
				String schema = DB.getDatabase().getSchema();
				String tableName = table.getTableName();
				if (md.storesUpperCaseIdentifiers())
				{
					tableName = tableName.toUpperCase();
				}
				else if (md.storesLowerCaseIdentifiers())
				{
					tableName = tableName.toLowerCase();
				}
				int noColumns = 0;
				String sql = null;
				//
				rs = md.getColumns(catalog, schema, tableName, null);
				while (rs.next())
				{
					noColumns++;
					String columnName = rs.getString ("COLUMN_NAME");
					if (!columnName.equalsIgnoreCase(column.getColumnName()))
						continue;
					
					//	update existing column
					boolean notNull = DatabaseMetaData.columnNoNulls == rs.getInt("NULLABLE");
					sql = column.getSQLModify(table, column.isMandatory() != notNull);
					if (DB.isOracle()) {
						// IDEMPIERE-3842 problem with oracle alter CLOB or BLOB
						int actualType = rs.getInt("DATA_TYPE");
						if (actualType == Types.CLOB) {
							if (sql.contains(" MODIFY " + column.getColumnName() + " CLOB")) {
								// trying to make CLOB a column that is already a CLOB
								sql = sql.replaceFirst(" MODIFY " + column.getColumnName() + " CLOB", " MODIFY " + column.getColumnName());
							}
						} else if (actualType == Types.BLOB) {
							if (sql.contains(" MODIFY " + column.getColumnName() + " BLOB")) {
								// trying to make BLOB a column that is already a BLOB
								sql = sql.replaceFirst(" MODIFY " + column.getColumnName() + " BLOB", " MODIFY " + column.getColumnName());
							}
						}
					}
					break;
				}
				DB.close(rs);
				rs = null;
			
				boolean isNoTable = noColumns == 0;
				//	No Table
				if (isNoTable)
					sql = table.getSQLCreate ();
				//	No existing column
				else if (sql == null)
					sql = column.getSQLAdd(table);
				
				if (isNoTable)
				{
					MColumn[] cols = table.getColumns(false);
					for (MColumn col : cols)
					{
						String fkConstraintSql = MColumn.getForeignKeyConstraintSql(md, catalog, schema, tableName, table, col, false);
						if (fkConstraintSql != null && fkConstraintSql.length() > 0)
							sql += fkConstraintSql;
					}
				}
				else
				{
					String fkConstraintSql = MColumn.getForeignKeyConstraintSql(md, catalog, schema, tableName, table, column, false);
					if (fkConstraintSql != null && fkConstraintSql.length() > 0)
						sql += fkConstraintSql;
				}
				
				int no = 0;
				if (sql.indexOf(DB.SQLSTATEMENT_SEPARATOR) == -1)
				{
					no = DB.executeUpdate(sql, false, null); 
				}
				else
				{
					String statements[] = sql.split(DB.SQLSTATEMENT_SEPARATOR);
					for (int i = 0; i < statements.length; i++)
					{
						int count = DB.executeUpdateEx(statements[i],null); 
						no += count;
					}
				}
		
				if (no == -1)
				{
					StringBuilder msg = new StringBuilder("@Error@ ");
					ValueNamePair pp = CLogger.retrieveError();
					if (pp != null)
						msg = new StringBuilder(pp.getName()).append(" - ");
					msg.append(sql);
					throw new AdempiereUserError (msg.toString());
				}
				allSql.append(sql);
				if (isNoTable)
					break;
			} finally {
				DB.close(rs);
				rs = null;
				if (conn != null) {
					try {
						conn.close();
					} catch (Exception e) {}
				}
			}
		  }
		} 
	  return tablesSynched;
	}

	public String removexSpaces(String snip) {
		String done = snip.replaceAll("\\s+","");
		return done;
	}

	private int createReference(HSSFCell cellIn){ 
		if (cellIn==null)
			throw new AdempiereException("Reference List has NULL at COLUMN: "+cellIn.getColumnIndex()+1);
		String colname = cellIn.getStringCellValue();
		reference = 17;
		X_AD_Reference ref = new Query(Env.getCtx(),X_AD_Reference.Table_Name,X_AD_Reference.COLUMNNAME_Name+"=?", trxName)
		.setParameters(colname)
		.first();
		if (ref==null) {//create new reference if not existed.
			ref = new X_AD_Reference(Env.getCtx(),0,trxName);
			ref.setValidationType(X_AD_Reference.VALIDATIONTYPE_ListValidation);
			ref.setEntityType(entity);
			ref.setName(colname);
			ref.saveEx(trxName);
		}  
		return ref.get_ID();
	} 
	
	void createRefList(HSSFCell cellIn, int y) {
		if (cellIn==null)
			return;
		String name = cellIn.getStringCellValue();
		if (name.isBlank())
			return;
		MRefList reflist = new Query(Env.getCtx(),MRefList.Table_Name,MRefList.COLUMNNAME_Name+"=? AND "
				+MRefList.COLUMNNAME_AD_Reference_ID+"=?", trxName)
				.setParameters(name,refset.get(y))
				.first();
		if (reflist==null) { // only new value shall be added
			MRefList reflistitem = new MRefList(Env.getCtx(),0,trxName); 
			reflistitem.setAD_Reference_ID(refset.get(y));
			reflistitem.setName(name);
			reflistitem.setValue(name.substring(0,name.length()<2?1:2));
			reflistitem.saveEx(trxName);
			listvaluesadded++;
		}		
	}
 
	 String getSQLCreate(MTable table)
	{
		String tableName=table.getTableName();
		StringBuffer sb = new StringBuffer("CREATE TABLE ")
			.append(tableName).append(" (");
		//
		boolean hasPK = false;
		boolean hasParents = false;
		StringBuffer constraints = new StringBuffer();
		MColumn[]	m_columns = null;
		m_columns = table.getColumns(true);
		for (int i = 0; i < m_columns.length; i++)
		{
			MColumn column = m_columns[i];
			String colSQL = column.getSQLDDL();
			if ( colSQL != null )
			{
				if (i > 0)
					sb.append(", ");
					sb.append(column.getSQLDDL());
			}
			else // virtual column
				continue;
			//
			if (column.isKey())
				hasPK = true;
			if (column.isParent())
				hasParents = true;
			String constraint = column.getConstraint(tableName);
			if (constraint != null && constraint.length() > 0)
				constraints.append(", ").append(constraint);
		}
		//	Multi Column PK
		if (!hasPK && hasParents)
		{
			StringBuffer cols = new StringBuffer();
			for (int i = 0; i < m_columns.length; i++)
			{
				MColumn column = m_columns[i];
				if (!column.isParent())
					continue;
				if (cols.length() > 0)
					cols.append(", ");
				cols.append(column.getColumnName());
			}
			
			if (tableName.toUpperCase().endsWith("_TRL"))
			{
				String constraintName = "";
				if (tableName.length()>22)
					constraintName=tableName.replace("_", "") + "Key";
				else
					constraintName = tableName + "_Key";
				
				sb.append(", CONSTRAINT ")
					.append(constraintName).append(" PRIMARY KEY (")
					.append(cols).append(")");				
			}
			else //(=standard)
			{
			sb.append(", CONSTRAINT ")
				.append(tableName).append("_Key PRIMARY KEY (")
				.append(cols).append(")");
			}
		}
		
		sb.append(constraints)
			.append(")");
		return sb.toString();
	}	//	getSQLCreate


	/**
	 * Taken from base plugin > org.compiere.model.Callout_AD_Column
	 * @param column
	 */
	public MColumn CalloutProcessor(MColumn column){ 
			// get defaults from most used case
			String sql = ""
					+ "SELECT AD_Reference_ID, "
					+ "       AD_Val_Rule_ID, "
					+ "       AD_Reference_Value_ID, "
					+ "       FieldLength, "
					+ "       DefaultValue, "
					+ "       MandatoryLogic, "
					+ "       ReadOnlyLogic, "
					+ "       IsIdentifier, "
					+ "       IsUpdateable, "
					+ "       IsAlwaysUpdateable, "
					+ "       FKConstraintType,"	
					+ "       COUNT(*) "
					+ "FROM   AD_Column "
					+ "WHERE  ColumnName = ? "
					+ "GROUP  BY AD_Reference_ID, "
					+ "          AD_Val_Rule_ID, "
					+ "          AD_Reference_Value_ID, "
					+ "          FieldLength, "
					+ "          DefaultValue, "
					+ "          MandatoryLogic, "
					+ "          ReadOnlyLogic, "
					+ "          IsIdentifier, "
					+ "          IsUpdateable, "
					+ "          IsAlwaysUpdateable, "
					+ "          FKConstraintType "
					+ "ORDER  BY COUNT(*) DESC ";
			PreparedStatement pstmt = null;
			ResultSet rs = null;
			try
			{
				pstmt = DB.prepareStatement(sql, null);
				pstmt.setString(1, column.getColumnName());
				rs = pstmt.executeQuery();
				if (rs.next()) {
					int ad_reference_id = rs.getInt(1);
					if (ad_reference_id == DisplayType.ID)
						ad_reference_id = DisplayType.TableDir;
					column.setAD_Reference_ID(ad_reference_id);
					column.setAD_Val_Rule_ID(rs.getInt(2));
					column.setAD_Reference_Value_ID(rs.getInt(3));
					column.setFieldLength(rs.getInt(4));
					column.setDefaultValue(rs.getString(5));
					column.setMandatoryLogic(rs.getString(6));
					column.setReadOnlyLogic(rs.getString(7));
					column.setIsIdentifier("Y".equals(rs.getString(8)));
					column.setIsUpdateable("Y".equals(rs.getString(9)));
					column.setIsAlwaysUpdateable("Y".equals(rs.getString(10)));
					column.setFKConstraintType(rs.getString(11));
				}
			}
			catch (SQLException e)
			{
				throw new DBException(e);
			}
			finally
			{
				DB.close(rs, pstmt);
				rs = null;
				pstmt = null;
			}
			return column;
		}
	
	/**
	 * 	First fetch parent's tab 
		Placing parent ID as ForeignKey
		Link to parent column if > column specified
	 * @param subTabName
	 * @param parentLinkColumnName
	 * @param tablepo
	 * @return
	 * @throws Exception 
	 */
	private int setForeignKey(String tablename, String parentLinkColumnName, MTable tablepo, int y) throws Exception {				 
		int parentLinkID = 0;
		String columnname = prefix+tablename.replaceAll("\\s+","")+"_ID"; 
		if (!parentLinkColumnName.isEmpty()){ 
			MColumn linkcolumn = tablepo.getColumn(prefix+parentLinkColumnName.replaceAll("\\s+","")+"_ID");
			if (linkcolumn!=null){
				columnname = parentLinkColumnName;
				parentLinkID = linkcolumn.get_ID();
			}
		} 
		MColumn columnexisted = tablepo.getColumn(columnname);
		if (columnexisted==null) {
			addColumn(y,columnname);			
			//fetch table model again after the new column add
			MTable check = new Query(Env.getCtx(), MTable.Table_Name,MTable.COLUMNNAME_TableName+"=?",trxName)
				.setParameters(tablepo.getTableName())
				.first();
			//check that this parent column exactly exist after all this
			MColumn column = check.getColumn(columnname);
			if (column==null) {  
				return 0;
				}
			column.setIsParent(true); 
			column.saveEx(trxName);
		}else {
			columnexisted.setIsParent(true); 
			columnexisted.saveEx(trxName);
		}

		return parentLinkID;
	}
	
	public static String leadingCapsSpacing(String subName) {
		String[] r = subName.split("(?=\\p{Lu})");
		StringBuilder composed = new StringBuilder();
		if (r.length>1){
			for (String s:r){
				s = s.trim();
				if (s.isEmpty()) 
					continue;
				if (s.length()==1)
					composed.append(s);
				else
					composed.append(s+" ");
			}
			return composed.toString().trim(); 
		}
		return subName;
	}
	private void workbookWrite() throws IOException {
		 
		FileOutputStream out = new FileOutputStream(File_Directory);
		if(out!=null)
		{
			workbook.write(out);
			out.close();
		}
	}
	
	void workbookClose() throws IOException { 
		workbook.close();
	}
}
