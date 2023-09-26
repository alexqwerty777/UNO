import pandas as pd

# df = pd.read_excel('Uno_.xlsx')

xlsx_file = pd.ExcelFile('Uno_.xlsx')

sheet_names = [name for name in xlsx_file.sheet_names if 'Матч' in name]
print(sheet_names)


def get_sheet_names(filename):
    return [name for name in pd.ExcelFile(filename).sheet_names if 'Матч' in name]


def count_results(filename: str, sheets: list):
    game_played = {}
    loosers = {}
    loose_percent = {}
    round_played = {}
    win_number = {}
    win_percent = {}
    point_numbers = {}
    avg_round_points = {}
    for game in sheets:
        df = pd.read_excel(filename, sheet_name=game)
        df = df.drop(index=[0, 1, 2])
        df = df.dropna()
        df = df.iloc[:, 1:]
        looser = df.sum().idxmax()
        loosers[looser] = loosers.get(looser) + 1 if loosers.get(looser) else 1

        for name in df.columns:
            game_played[name] = game_played.get(name) + 1 if game_played.get(name) else 1
            round_played[name] = round_played.get(name) + len(df.index) if round_played.get(name) else len(df.index)
            print(df[name].value_counts())
            wins = df[name].value_counts()[0.0] if 0.0 in df[name].value_counts().keys() else 0
            win_number[name] = win_number.get(name) + wins if win_number.get(name) else wins
            point_numbers[name] = point_numbers.get(name) + int(df[name].sum()) if point_numbers.get(name) else int(
                df[name].sum())
            # point_numbers = df[name].sum()

            print(f'{name}, add {len(df.index)}, and become {game_played.get(name)}')

    for name, game_number in game_played.items():
        loose_percent[name] = 0 if not loosers.get(name) else round(loosers.get(name) / game_number)

    win_percent = {key: round(win_number.get(key) / round_played.get(key), 2) for key in win_number.keys()}

    avg_round_points = {key: round(point_numbers.get(key) / round_played.get(key), 2) for key in win_number.keys()}

    print(f'Loosers: {loosers}')
    print(f'Game played: {game_played}')
    print(f'Loose percent: {loose_percent}')
    print(f'Round played: {round_played}')
    print(f'Win number: {win_number}')
    print(f'Win percent: {win_percent}')
    print(f'Point number: {point_numbers}')
    print(f'AVG round points: {avg_round_points}')
    data = [
            game_played, loosers, loose_percent, round_played,
            win_number, win_percent, point_numbers, avg_round_points
            ]
    new_df = pd.DataFrame(data)
    new_df.fillna(value=0, inplace=True)
    print(new_df)
    return new_df


# def get_screenshoot():
#     screen = pyautogui.screenshot('screensot.png')
#     print(screen)


def main():
    filename = 'Uno_.xlsx'
    # get_screenshoot()
    sheets = get_sheet_names(filename)
    df = count_results(filename, sheets)
    df.to_excel('Results.xlsx')


if __name__ == '__main__':
    main()
