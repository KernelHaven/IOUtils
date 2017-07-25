package net.ssehub.kernel_haven.io;

import java.io.File;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;
import org.junit.runners.Suite.SuiteClasses;

import net.ssehub.kernel_haven.io.csv.AllCSVTests;

/**
 * Test suite for the whole plug-in.
 * @author El-Sharkawy
 *
 */
@RunWith(Suite.class)
@SuiteClasses({AllCSVTests.class})
public class AllTests {
    
    public static final File TESTDATA = new File("testdata");

}
