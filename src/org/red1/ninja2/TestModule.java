package org.red1.ninja2;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import org.adempiere.exceptions.AdempiereException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.compiere.Adempiere;
import org.compiere.model.MColumn;
import org.compiere.model.MTable;
import org.compiere.util.CLogger;
import org.compiere.util.DB;
import org.compiere.util.Env;
import org.compiere.util.Ini;
import org.compiere.util.Trx;
import org.eclipse.equinox.app.IApplication;
import org.eclipse.equinox.app.IApplicationContext;

public class TestModule implements IApplication {
    private static final CLogger log = CLogger.getCLogger(TestModule.class);
    private String File_Directory = "/home/red1/hongkong/Ninja.xls";
    private int p_ValueLength = 22;
    private int p_NameLength = 22;
    private HSSFWorkbook workbook = null;
    private int reference = 10;
    private List<Integer> refset = new ArrayList<>();
    private String trxName = "";
    private int ref_cnt = 0;
    private int model_cnt = 0;
    private int listvaluesadded = 0;
    private String prefix = "CA_";
    private String entity = "A";
    private String ad_name;
    private List<String> masterset = new ArrayList<>();
    private List<MTable> tableset = new ArrayList<>();
    // ... other fields ...

    @Override
    public Object start(IApplicationContext context) throws Exception {
        initializeEnvironment();
        processWorkbook();
        return IApplication.EXIT_OK;
    }

    @Override
    public void stop() {
        cleanupResources();
    }

    public static void main(String[] args) throws Exception {
        ModuleMaker maker = new ModuleMaker();
        maker.startStandalone(args);
    }

    public void startStandalone(String[] args) throws Exception {
        initializeEnvironment();
        processWorkbook();
        cleanupResources();
    }

    private void initializeEnvironment() throws Exception {
        if (!Adempiere.isStarted()) {
            String propertyFile = Ini.getFileName(false);
            File file = new File(propertyFile);
            if (!file.exists()) {
                throw new AdempiereException("Property file not found: " + file.getAbsolutePath());
            }
            Adempiere.startup(false);
        }

        trxName = "Ninja_" + System.currentTimeMillis();
        Env.setContext(Env.getCtx(), "#AD_Client_ID", 0);
        Env.setContext(Env.getCtx(), "#AD_Org_ID", 0);
    }

    private void processWorkbook() throws Exception {
        try (FileInputStream file = new FileInputStream(File_Directory)) {
            workbook = new HSSFWorkbook(file);
            log.info(workbook.getNumberOfSheets() + " < worksheets in " + File_Directory);

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                HSSFSheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                if ("List".equals(sheetName)) {
                    doList(sheet);
                } else if ("Model".equals(sheetName)) {
                    doModel(sheet);
                } else if (sheetName.split("_").length > 1) {
                    importModel(sheet);
                }
            }
        }
    }

    private void cleanupResources() {
        try {
            if (workbook != null) {
                workbook.close();
            }
            if (trxName != null && !trxName.isEmpty()) {
                Trx trx = Trx.get(trxName, false);
                if (trx != null) {
                    trx.close();
                }
            }
        } catch (IOException e) {
            log.severe("Error closing workbook: " + e.getMessage());
        }
    }

    // ... keep your existing methods like doList, doModel, importModel, etc. ...

    // Modify other methods to use log instead of System.out.println
    // For example:
    private void doList(HSSFSheet sheet) {
        // ... existing code ...
        log.info(ref_cnt + " < Reference Table Count --- SYSTEM COMPLETED --- TRANSACTION NAME = " + trxName);
        log.info(listvaluesadded + " < ReferenceList Table values added");
    }

    // ... other methods ...
}