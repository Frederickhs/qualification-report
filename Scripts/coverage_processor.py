from qualification_reader import read_all_qualifications


VENUE_MODULES = {
    "Unchained": [
        "UCMOD0",
        "UCMOD1",
        "UCMOD2",
        "UCMOD3",
        "UCMOD4",
    ],
    "Wolf": [
        "CUMOD0",
        "CUMOD1",
        "CUMOD2",
        "CUMOD3",
    ],
    "Darkmoor": [
        "DUMOD0",
        "DUMOD1",
    ],
}

REPORT_GROUPS = {
    "Operational": {
        "schedule_groups": {
            "Operational - 1st Shift",
            "Operational - 2nd Shift",
        },
        "venues": {
            "Unchained",
            "Wolf",
            "Darkmoor",
        },
    },
    "3rd Shift Unchained": {
        "schedule_groups": {
            "3rd Shift Unchained",
        },
        "venues": {
            "Unchained",
        },
    },
    "3rd Shift Wolf": {
        "schedule_groups": {
            "3rd Shift Wolf",
        },
        "venues": {
            "Wolf",
            "Darkmoor",
        },
    },
}

TOTAL_GROUPS = {
    "Unchained": {
        "Operational - 1st Shift",
        "Operational - 2nd Shift",
        "3rd Shift Unchained",
    },
    "Wolf": {
        "Operational - 1st Shift",
        "Operational - 2nd Shift",
        "3rd Shift Wolf",
    },
    "Darkmoor": {
        "Operational - 1st Shift",
        "Operational - 2nd Shift",
        "3rd Shift Wolf",
    },
}


def get_population(technicians, schedule_groups):
    return [
        technician
        for technician in technicians
        if technician["schedule_group"] in schedule_groups
    ]


def calculate_module_result(
    population,
    qualification_records,
    qualification_code,
):
    population_ssos = {
        technician["sso"]
        for technician in population
    }

    qualified_ssos = {
        sso
        for sso, qualification
        in qualification_records
        if qualification == qualification_code
        and sso in population_ssos
    }

    population_count = len(population_ssos)
    qualified_count = len(qualified_ssos)

    percentage = (
        qualified_count / population_count
        if population_count
        else 0
    )

    return {
        "qualification": qualification_code,
        "qualified_count": qualified_count,
        "population_count": population_count,
        "percentage": percentage,
        "qualified_ssos": qualified_ssos,
    }


def calculate_venue_result(
    population,
    qualification_records,
    venue,
):
    module_results = []

    for qualification_code in VENUE_MODULES[venue]:
        module_result = calculate_module_result(
            population=population,
            qualification_records=qualification_records,
            qualification_code=qualification_code,
        )

        module_results.append(module_result)

    population_count = len(
        {
            technician["sso"]
            for technician in population
        }
    )

    total_qualified_occurrences = sum(
        module_result["qualified_count"]
        for module_result in module_results
    )

    possible_qualification_occurrences = (
        population_count * len(VENUE_MODULES[venue])
    )

    total_percentage = (
        total_qualified_occurrences
        / possible_qualification_occurrences
        if possible_qualification_occurrences
        else 0
    )

    final_module = VENUE_MODULES[venue][-1]

    final_module_result = next(
        result
        for result in module_results
        if result["qualification"] == final_module
    )

    return {
        "venue": venue,
        "population_count": population_count,
        "module_results": module_results,
        "total_qualified_occurrences": (
            total_qualified_occurrences
        ),
        "possible_qualification_occurrences": (
            possible_qualification_occurrences
        ),
        "total_percentage": total_percentage,
        "fully_qualified_count": (
            final_module_result["qualified_count"]
        ),
        "fully_qualified_percentage": (
            final_module_result["percentage"]
        ),
    }


def calculate_dashboard_results():
    qualification_data = read_all_qualifications()

    technicians = qualification_data["technicians"]
    qualification_records = qualification_data[
        "qualification_records"
    ]

    dashboard_results = {
        "sections": {},
        "totals": {},
        "exceptions": {
            "line_schedule": qualification_data[
                "line_schedule_exceptions"
            ],
            "qualification_files": qualification_data[
                "file_exceptions"
            ],
            "inactive_records": qualification_data[
                "inactive_records"
            ],
            "invalid_dates": qualification_data[
                "invalid_date_records"
            ],
            "duplicate_records": qualification_data[
                "duplicate_records"
            ],
        },
    }

    for section_name, section_rules in REPORT_GROUPS.items():
        population = get_population(
            technicians=technicians,
            schedule_groups=section_rules[
                "schedule_groups"
            ],
        )

        section_result = {
            "population_count": len(population),
            "venues": {},
        }

        for venue in VENUE_MODULES:
            if venue not in section_rules["venues"]:
                continue

            section_result["venues"][venue] = (
                calculate_venue_result(
                    population=population,
                    qualification_records=(
                        qualification_records
                    ),
                    venue=venue,
                )
            )

        dashboard_results["sections"][
            section_name
        ] = section_result

    for venue, schedule_groups in TOTAL_GROUPS.items():
        population = get_population(
            technicians=technicians,
            schedule_groups=schedule_groups,
        )

        dashboard_results["totals"][venue] = (
            calculate_venue_result(
                population=population,
                qualification_records=(
                    qualification_records
                ),
                venue=venue,
            )
        )

    return dashboard_results


def format_percentage(value):
    return f"{value:.1%}"


def display_venue_result(venue_result):
    print(
        f"{venue_result['venue']} "
        f"(Population: "
        f"{venue_result['population_count']})"
    )

    for module_result in venue_result["module_results"]:
        print(
            f"  {module_result['qualification']}: "
            f"{module_result['qualified_count']} of "
            f"{module_result['population_count']} | "
            f"{format_percentage(module_result['percentage'])}"
        )

    print(
        "  Venue Total: "
        f"{venue_result['total_qualified_occurrences']} of "
        f"{venue_result['possible_qualification_occurrences']} | "
        f"{format_percentage(venue_result['total_percentage'])}"
    )

    print(
        "  Fully Qualified: "
        f"{venue_result['fully_qualified_count']} of "
        f"{venue_result['population_count']} | "
        f"{format_percentage(
            venue_result['fully_qualified_percentage']
        )}"
    )


def display_dashboard_results():
    results = calculate_dashboard_results()

    print("=" * 80)
    print("QUALIFICATION COVERAGE RESULTS")
    print("=" * 80)
    print()

    for section_name, section_result in (
        results["sections"].items()
    ):
        print("-" * 80)
        print(
            f"{section_name} "
            f"(Roster Population: "
            f"{section_result['population_count']})"
        )
        print("-" * 80)

        for venue_result in (
            section_result["venues"].values()
        ):
            display_venue_result(venue_result)
            print()

    print("-" * 80)
    print("TOTAL")
    print("-" * 80)

    for venue_result in results["totals"].values():
        display_venue_result(venue_result)
        print()

    exception_count = sum(
        len(exception_list)
        for exception_list
        in results["exceptions"].values()
    )

    print("-" * 80)
    print(f"Total exceptions: {exception_count}")

    for exception_type, exception_list in (
        results["exceptions"].items()
    ):
        print(
            f"  {exception_type}: "
            f"{len(exception_list)}"
        )


if __name__ == "__main__":
    display_dashboard_results()