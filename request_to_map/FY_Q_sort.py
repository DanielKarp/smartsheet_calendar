from datetime import date, datetime

YEARS_AND_QUARTERS = {
    20: {
        1: (date(2019, 7, 28), date(2019, 10, 26)),
        2: (date(2019, 10, 27), date(2020, 1, 25)),
        3: (date(2020, 1, 26), date(2020, 4, 25)),
        4: (date(2020, 4, 26), date(2020, 7, 25)),
    },
    21: {
        1: (date(2020, 7, 26), date(2020, 10, 24)),
        2: (date(2020, 10, 25), date(2021, 1, 23)),
        3: (date(2021, 1, 24), date(2021, 5, 1)),
        4: (date(2021, 5, 2), date(2021, 7, 31)),
    },
    22: {
        1: (date(2021, 8, 1), date(2021, 10, 24)),
        2: (date(2021, 10, 25), date(2022, 1, 23)),
        3: (date(2022, 1, 24), date(2022, 5, 1)),
        4: (date(2022, 5, 2), date(2022, 7, 31)),
    },
    23: {
        1: (date(2022, 8, 1), date(2022, 10, 29)),
        2: (date(2022, 10, 30), date(2023, 1, 28)),
        3: (date(2023, 1, 29), date(2023, 4, 29)),
        4: (date(2023, 4, 30), date(2023, 7, 29)),
    },
    24: {
        1: (date(2023, 7, 30), date(2023, 10, 28)),
        2: (date(2023, 10, 29), date(2024, 1, 27)),
        3: (date(2024, 1, 28), date(2024, 4, 27)),
        4: (date(2024, 4, 28), date(2024, 7, 27)),
    },
    25: {
        1: (date(2024, 7, 28), date(2024, 10, 27)),
        2: (date(2024, 10, 28), date(2025, 2, 2)),
        3: (date(2025, 2, 3), date(2025, 4, 28)),
        4: (date(2025, 4, 29), date(2025, 7, 26)),
    },
    26: {
        1: (date(2025, 7, 27), date(2025, 10, 25)),
        2: (date(2025, 10, 26), date(2026, 1, 24)),
        3: (date(2026, 1, 25), date(2026, 4, 25)),
        4: (date(2026, 4, 26), date(2026, 7, 25)),
    },
    27: {  # estimated, need to update once calendar is out
        1: (date(2025, 7, 26), date(2025, 10, 25)),
        2: (date(2025, 10, 26), date(2026, 1, 24)),
        3: (date(2026, 1, 25), date(2026, 4, 25)),
        4: (date(2026, 4, 26), date(2026, 7, 25)),
    },
}


def calc_fy_q_hardcoded(input_date):
    input_date = datetime.strptime(input_date, '%Y-%m-%d').date()
    for year, quarters in YEARS_AND_QUARTERS.items():
        for quarter, time_range in quarters.items():
            if time_range[0] <= input_date <= time_range[1]:
                return year, quarter
    else:
        # if all else fails, return the next year after the last known year, Q1
        return max(YEARS_AND_QUARTERS.keys()) + 1, 1


def print_quarters():
    for year, quarters in YEARS_AND_QUARTERS.items():
        print(f'FY{year}:')
        for quarter, time_range in quarters.items():
            print(f'\tQ{quarter}:')
            print(f'\t\ttime range: {time_range[0]} – {time_range[1]}')
            print(f'\t\tduration: {(time_range[1] - time_range[0]).days}')


if __name__ == '__main__':
    test_dates = ['2017-04-20',
                  '2020-01-01',
                  '2020-07-22',
                  '2020-07-27',
                  '2021-07-31',
                  '2021-10-08',
                  '2022-10-08',
                  '2023-10-08',
                  '2024-10-08', ]
    print('test:')
    for test_date in test_dates:
        print(test_date)
        # second test - calculate quarter from a hardcoded list of quarters with start/end dates
        result = calc_fy_q_hardcoded(test_date)
        print('result:')
        print(f'\t{result}')
        print(f'\tFY{result[0]} Q{result[1]}\n')

    # just print all the years and quarters
    print_quarters()
