import ExcelHandler.ExcelHandler;
import Model.Volunteer;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * downloadRoot must be changed in order to test the download method.
 */

public class ExcelHandlerShould {

    List<Volunteer> volunteerList;

    ExcelHandler handler;

    String downloadRoot = "C:/Users/Jesús Artiles/Desktop/testdata.xlsx";

    String[] CORRECT_DATA_EXPECTED = {"NAME EMAIL","Jesus test_email_1@gmail.com","Paco test_email_2@gmail.com",
            "Juanma test_email_3@gmail.com"};

    String[] CORRECT_DATA_EXPECTED_AFTER_ADD = {"NAME EMAIL","Jesus test_email_1@gmail.com",
            "Paco test_email_2@gmail.com", "Juanma test_email_3@gmail.com", "Messi test_email_4@gmail.com"};

    String[] CORRECT_DATA_EXPECTED_AFTER_DELETE = {"NAME EMAIL","Jesus test_email_1@gmail.com",
            "Juanma test_email_3@gmail.com"};

    String[] CORRECT_DATA_EXPECTED_AFTER_CLEAR = {};

    @Before
    public void setup() throws Exception {
        volunteerList = new ArrayList<>();
        volunteerList.add(new Volunteer("Jesus","test_email_1@gmail.com"));
        volunteerList.add(new Volunteer("Paco","test_email_2@gmail.com"));
        volunteerList.add(new Volunteer("Juanma","test_email_3@gmail.com"));

        handler = new ExcelHandler(volunteerList, "./testdata.xlsx");
        handler.initiate();
        handler.clearExcel();
    }

    @Test
    public void write_volunteers_in_excel_correctly() throws Exception{
        handler.writeVolunteersInExcel();

        assertThat(handler.readExcel()).isEqualTo(CORRECT_DATA_EXPECTED);
    }

    @Test
    public void clear_excel_correctly() throws IOException {
        handler.writeVolunteersInExcel();

        handler.clearExcel();

        assertThat(handler.readExcel()).isEqualTo(CORRECT_DATA_EXPECTED_AFTER_CLEAR);
    }

    @Test
    public void add_new_volunteer_correctly() throws Exception {
        handler.writeVolunteersInExcel();

        handler.addVolunteer(new Volunteer("Messi","test_email_4@gmail.com"));

        assertThat(handler.readExcel()).isEqualTo(CORRECT_DATA_EXPECTED_AFTER_ADD);
    }

    @Test
    public void delete_volunteer_correctly() throws Exception {
        handler.writeVolunteersInExcel();

        handler.deleteVolunteer(new Volunteer("Paco","test_email_2@gmail.com"));

        assertThat(handler.readExcel()).isEqualTo(CORRECT_DATA_EXPECTED_AFTER_DELETE);
    }

    @Test
    public void download_excel_correctly() throws IOException {
        handler.writeVolunteersInExcel();

        handler.downloadExcel(downloadRoot);

        File excel = new File(downloadRoot);

        assertThat(excel.exists()).isTrue();
    }
}
