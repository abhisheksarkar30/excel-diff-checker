package edu.abhi.poi.excel;

import static org.junit.Assert.*;

import org.junit.Test;
import org.junit.Before;

public class CallableValueTest {

    private CallableValue cv;

    @Before
    public void setup(){
        cv = new CallableValue();
    }

    @Test
    public void setFailedFalseTest(){
        cv.setFailed(false);
        assertFalse(cv.isFailed());
    }

    @Test
    public void setFailedTrueTest(){
        cv.setFailed(true);
        assertTrue(cv.isFailed());
    }

    @Test
    public void setExceptionTest(){
        Exception e = new RuntimeException("This is an error!");
        cv.setException(e);
        assertEquals(e, cv.getException());
    }
}
