package de.ewoelfel.caretool;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Component;

import java.util.Arrays;
import java.util.Map;

import static java.util.stream.Collectors.toMap;

@Component
public class CmdRunner implements ApplicationRunner {

    @Autowired
    private DocumentHandler documentHandler;

    private static final Logger logger = LoggerFactory.getLogger(CmdRunner.class);

    @Override
    public void run(ApplicationArguments args) throws Exception {
        logger.info("Caretool started:");
        Map<String, String> optionMap = args.getOptionNames().stream()
                .collect(toMap(name -> name, name -> args.getOptionValues(name).get(0)));

        GenerationContext context = new GenerationContext(optionMap);
        //create document based on template
        String documentName = documentHandler.generateDocument(context);
        logger.info("done generating "+ documentName);
        System.exit(1);
    }
}
