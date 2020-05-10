package edu.abhi.poi.excel;

import java.util.List;
import java.util.concurrent.Future;

import org.apache.commons.collections4.CollectionUtils;

/**
 * 
 * @author abhishek sarkar
 *
 */
public class CallableValue {
	
	/** Failure flag */
	private boolean failed = false;
	
	/** Diff Flag */
	private boolean diffFlag = false;

	/** The exception on failure */
	private Exception exception = null;
	
	/** To contain the diff report */
	private StringBuilder diffContainer = new StringBuilder();

	/**
	 * @return the failed
	 */
	public boolean isFailed() {
		return failed;
	}

	/**
	 * @param failed
	 *            the failed to set
	 */
	public void setFailed(boolean failed) {
		this.failed = failed;
	}

	/**
	 * @return the exception
	 */
	public Exception getException() {
		return exception;
	}

	/**
	 * @param exception
	 *            the exception to set
	 */
	public void setException(Exception exception) {
		this.exception = exception;
	}
	
	/**
	 * Analyze the output of the threads. If any of the thread failed, it
	 * prints the Exception stack trace.
	 * 
	 * @param futures
	 *            the list of {@link Future} which contains the returned
	 *            {@link CallableValue} for all the threads.
	 * @return if any diff found
	 */
	public static boolean analyzeResult(List<Future<CallableValue>> futures) {
		boolean diffFound = false;
		if (!CollectionUtils.isEmpty(futures)) {
			try {
				for (Future<CallableValue> future : futures) {
					CallableValue value = future.get();
					if(value != null && value.isDiffFlag()) {
						System.out.println(value.getDiffContainer());		
						diffFound |= value.isDiffFlag();
					}
					if (value != null && value.isFailed()) {
						value.getException().printStackTrace(System.out);
					}
				}
			} catch (Exception e) {
				e.printStackTrace(System.out);
			}
		}
		return diffFound;
	}

	/**
	 * @return the diffFlag
	 */
	public boolean isDiffFlag() {
		return diffFlag;
	}

	/**
	 * @param diffFlag the diffFlag to set
	 */
	public void setDiffFlag(boolean diffFlag) {
		this.diffFlag = diffFlag;
	}

	/**
	 * @return the diffReport
	 */
	public StringBuilder getDiffContainer() {
		return diffContainer;
	}
}
