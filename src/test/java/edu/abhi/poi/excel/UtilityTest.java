package edu.abhi.poi.excel;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.IOException;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class UtilityTest {
    private File tempFile;
    @Before
    public void setup() throws IOException{
        tempFile = File.createTempFile("prefix-", "-suffix");
    }
    @After
    public void tearDown(){
        tempFile.delete();
    }
    @Test
    public void deleteIfExistsNonExistingFileTest(){
        assertTrue(tempFile.delete());
        boolean result = Utility.deleteIfExists(tempFile);
        assertFalse(result);
    }
     @Test
    public void deleteIfExistsExistingFileTest(){
        boolean result = Utility.deleteIfExists(tempFile);
        assertTrue(result);
    }
}
