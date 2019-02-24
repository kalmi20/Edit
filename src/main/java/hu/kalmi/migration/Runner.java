package hu.kalmi.migration;

import java.util.Set;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.stereotype.Component;

@Component
public class Runner implements ApplicationRunner {

	@Autowired
	private ExcelGenerator excelGenerator;
	
	@Autowired
	private MigrationTool migrationTool;
	
	@Override
	public void run(ApplicationArguments args) throws Exception {
		Set<String> optionNames = args.getOptionNames();
		if (optionNames.contains("migrate")) {
			migrationTool.run(args);
		} else {
			excelGenerator.run(args);
		}
	}
}
