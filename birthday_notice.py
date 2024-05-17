from zhdate import ZhDate
import datetime
import requests
import pandas as pd

# Specify the file path
file_path = r'./birthday.xlsx'

# Load the Excel file into a pandas DataFrame
df = pd.read_excel(file_path, sheet_name="Sheet1")

# Bark API endpoint
api_endpoint = 'https://api.day.app/xxxxxxxxxx/'

current_time = datetime.date.today()
current_year = datetime.date.today().year
lunar_date =  ZhDate.today().chinese()
lunar_today = ZhDate.today()
print(f'今日{current_time}\n{lunar_date}')
 
def main():
    for index, row in df.iterrows():
        try:
            name = row['Name']
            birth = row['Birthday']
            lunarbirth = row['LunarBirthday']
            year, month, day = birth.strftime('%Y-%m-%d').split('-')
            year = int(year)
            month = int(month)
            day = int(day)
            lunar_year, lunar_month, lunar_day = lunarbirth.strftime('%Y-%m-%d').split('-')
            lunar_month = int(lunar_month)
            lunar_day = int(lunar_day)
            how_old = current_year - year
            type = row['Type']
            # Check if the record is a solar birthday
            if type == "阳历":
                result = datetime.date(current_year, month, day)
                flag = (result - current_time).days
            # Check if the record is a lunar birthday
                print(f'{name}过阳历生日{birth},距今{flag}天')
            else:
                this_lunarbirth = ZhDate(current_year, lunar_month, lunar_day)
                result = ZhDate.to_datetime(this_lunarbirth).date()
                flag = this_lunarbirth - lunar_today
                print(f'{name}过阴历生日{lunarbirth},今年{this_lunarbirth},阳历是{result}，距今{flag}天')
            # Check if the birthday has passed or is within 3 days
            if flag < 0:
                continue
            elif flag <= 3:
                title = f'今日{current_time}\n{lunar_date}'
                if flag == 0:
                    desp = "{}今天过{}岁生日,阴历{}".format(name, how_old,lunarbirth.strftime('%m月%d日'))
                else:
                    desp = "{}{}({}天后)过{}岁生日,阴历{}".format(name, result.strftime('%Y-%m-%d'), flag, how_old,lunarbirth.strftime('%Y-%m-%d'))
                # Send POST request to the API endpoint with the notification text
                data = {'title': title, 'body': desp}
                response = requests.post(api_endpoint, data=data)
                # Check the response status
                if response.status_code == 200:
                    print('Notice sent successfully!')
                else:
                    print('Failed to send Notice.')
        except Exception as e:
            print(f"Error processing row {index + 1} for {name}: {e}")

if __name__ == '__main__':
    main()
