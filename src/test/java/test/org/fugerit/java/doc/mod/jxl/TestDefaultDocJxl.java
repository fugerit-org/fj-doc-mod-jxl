package test.org.fugerit.java.doc.mod.jxl;

import org.fugerit.java.doc.mod.jxl.XlsTypeHandler;
import org.junit.Assert;
import org.junit.Test;

public class TestDefaultDocJxl extends TestDocBase {

	@Test
	public void testOpenPDF() {
		boolean ok = this.testDocWorker( "default_doc" ,  XlsTypeHandler.HANDLER, false );
		Assert.assertTrue( ok );
	}
	
	@Test
	public void testFailPDF() {
		boolean ok = this.testDocWorker( "template_fail" ,  XlsTypeHandler.HANDLER, true );	
		Assert.assertTrue( ok );
	}
	
}
