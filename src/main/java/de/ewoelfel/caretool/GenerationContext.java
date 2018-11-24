package de.ewoelfel.caretool;

import de.jollyday.HolidayCalendar;
import de.jollyday.HolidayManager;
import lombok.Getter;

import java.time.LocalDate;
import java.time.Month;
import java.time.Year;
import java.time.ZoneId;
import java.time.temporal.ChronoUnit;
import java.util.*;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import static java.util.stream.Collectors.toList;
import static java.util.stream.Collectors.toMap;

@Getter
public class GenerationContext {

    private static final String OPTION_YEAR = "year";
    private static final String OPTION_NAME = "name";
    private static final String OPTION_COUNTRY = "country";
    private static final String OPTION_COUNTY = "county";

    private final Year currentYear = Year.now();
    private final Year nextYear = currentYear.plusYears(1);
    private final Year year;
    private final String name;
    private final Map<Month, List<Day>> daysInMonth = new LinkedHashMap<>();

    private HolidayManager holydayManager;

    /**
     * Combines the days to use for the shift schedule with holydays
     * @param optionMap
     */
    public GenerationContext(Map<String, String> optionMap) {
        this.year = Year.parse(optionMap.getOrDefault(OPTION_YEAR, String.valueOf(nextYear.getValue())));
        this.name = optionMap.getOrDefault(OPTION_NAME, "Unbekannt");

        //is used currently in mecklenburg vorpommerania
        String county = optionMap.getOrDefault(OPTION_COUNTY,"mv");

        holydayManager = HolidayManager.getInstance(HolidayCalendar.GERMANY);

        Stream.of(Month.values()).forEach(month -> {
            daysInMonth.put(month,
                    // length+1 as the range of intstream is exclusive
                    IntStream.range(1, month.length(year.isLeap()) + 1).boxed()
                            .map(v -> {
                                LocalDate date = LocalDate.of(year.getValue(), month, v);
                                GregorianCalendar cal = GregorianCalendar.from(date.atStartOfDay(ZoneId.systemDefault()));
                                return Day.builder()
                                        .value(date)
                                        .holyday(holydayManager.isHoliday(cal, county))
                                        .build();
                            })
                            .collect(toList())
            );
        });
    }
}
