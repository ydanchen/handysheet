import com.google.api.services.sheets.v4.Sheets;
import com.ydanchen.handysheet.SpreadSheet;
import com.ydanchen.handysheet.enums.Dimension;
import com.ydanchen.handysheet.enums.ValueInputOption;
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

        // Prepare some values
        Object[][] values = {
                {"A1", "B1", "C1"},
                {"A2", "B2", "C2"},
                {"A3", "B3", "C3"}
        };

        // Open a spreadsheet
        SpreadSheet spreadsheet = new SpreadSheet(service).withId(TEST_SHEET);

        // Write values on spreadsheet "Sheet1" in range "A1:C3"
        spreadsheet
                .onSheet("Sheet1")
                .toRange("A1:C3")
                .writeValues(values);

        // Append some value to the end of the spreadsheet
        Object[][] appendValues = {{"one", "two", "three"}};
        spreadsheet
                .onSheet("Sheet1")
                .toRange("A4:E4")
                .withValueInputOption(ValueInputOption.RAW)
                .appendValues(appendValues);

        // Insert one blank row at the top of spreadsheet
        // Starting from row #0 to row #1
        spreadsheet
                .onSheet("Sheet1")
                .select(Dimension.ROWS)
                .from(0)
                .to(1)
                .insertEmpty();

        // Delete row #2 (B)
        spreadsheet
                .onSheet("Sheet1")
                .select(Dimension.COLUMNS)
                .from(2)
                .to(3)
                .delete();
        // Merge range B2:D4
        spreadsheet
                .onSheet("Sheet1")
                .from(2,2)
                .to(4,4)
                .mergeCells();
    }
}