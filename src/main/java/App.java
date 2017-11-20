import com.google.api.services.sheets.v4.Sheets;
import com.ydanchen.handysheet.SpreadSheet;
import com.ydanchen.handysheet.services.SheetsServiceProvider;

import java.io.IOException;

/**
 * Handy SpreadSheet Demo App
 *
 * @author Yevhen Danchenko
 */
public class App {
    /**
     * Replace with your spreadsheet id
     * https://docs.google.com/spreadsheets/d/xxXXxXxXxxXXXXXxxxxXXXxXxxXxxxXXXxXxXXxxXxXX/edit
     */
    private static final String TEST_SHEET = "xxXXxXxXxxXXXXXxxxxXXXxXxxXxxxXXXxXxXXxxXxXX";

    public static void main(String[] args) throws IOException {
        /*
         * Create Sheets API client service
         * IMPORTANT! Make sure your sheet_client_secret.json file is present in the main/java/resources folder
         * See {@url https://developers.google.com/sheets/api/quickstart/java} for details
         */
        Sheets service = SheetsServiceProvider.createSheetsService("My Application");

        // Create a spreadsheet
        SpreadSheet sheet = new SpreadSheet.Builder(service)
                .withId(TEST_SHEET)
                .inRange("Sheet1!A1:A2")
                .build();

        // Insert one blank row at the top of sheet
        sheet.insertRows(0, 1, false);
    }
}