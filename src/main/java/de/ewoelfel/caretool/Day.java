package de.ewoelfel.caretool;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;

import java.time.LocalDate;

@Getter
@Builder
public class Day {

    private boolean holyday;
    private LocalDate value;
}
