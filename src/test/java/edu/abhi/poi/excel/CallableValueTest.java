package edu.abhi.poi.excel;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Future;

import org.junit.Before;
import org.junit.Test;

public class CallableValueTest {

    private CallableValue callableValue;

    @Before
    public void setup(){
        callableValue = new CallableValue();
    }

    @Test
    public void setExceptionTest(){
        Exception exception=new RuntimeException("Test Message");
        callableValue.setException(exception);
        assertEquals(exception,callableValue.getException());
    }

    @Test
    public void AnalyzeResultWithEmptyList(){
        List<Future<CallableValue>> futures = new ArrayList<>();
        assertFalse(CallableValue.analyzeResult(futures));

    }
    @Test
    public void setFailedTrueTest(){
        callableValue.setFailed(true);
        assertTrue(callableValue.isFailed());
    }
    @Test
    public void setFailedFalseTest(){
        callableValue.setFailed(false);
        assertFalse(callableValue.isFailed());
    }
    @Test
    public void getDiffContainerWithContentTest(){
        String differenceContent="Test difference";
        callableValue.getDiffContainer().append(differenceContent);

        assertEquals(differenceContent,callableValue.getDiffContainer().toString());
    }
    @Test
    public void getDiffContainerWithoutContentTest(){
        StringBuilder testStringBuilder=callableValue.getDiffContainer();
        assertTrue(testStringBuilder.isEmpty());
    }
    @Test
    public void setDiffFlagFalseTest(){
        callableValue.setDiffFlag(false);
        assertFalse(callableValue.isDiffFlag());
    }

    @Test
    public void setDiffFlagTrueTest(){
        callableValue.setDiffFlag(true);
        assertTrue(callableValue.isDiffFlag());
    }

}
