=IF(
    ISNA(MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0)),
    F37,
    IF(AND(F37="", INDEX(exer_old!F:F, MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0))=""),
        "",
        IF(F37="", INDEX(exer_old!F:F, MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0)),
            IF(INDEX(exer_old!F:F, MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0))="", F37,
                IF(INDEX(exer_old!F:F, MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0)) > F37,
                    INDEX(exer_old!F:F, MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0)),
                    F37
                )
            )
        )
    )
)



=IF(
  ISNA(MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
  I2,
  IF(
    AND(I2="", INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0))=""),
    I2,
    IF(
      INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0))="",
      I2,
      IF(
        I2="",
        INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
        IF(
          INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)) > I2,
          INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
          I2
        )
      )
    )
  )
)# Excelsheets
