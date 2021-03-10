import pandas as pd                        
from pytrends.request import TrendReq

def main():
    kw_and_dates = pd.read_excel("../kw_and_dates.xlsx")
    downLoadTrends(kw_and_dates)


def downLoadTrends(kw_and_dates):
    pytrends = TrendReq()
    for row in kw_and_dates.iterrows():
        kw_list = [row[1][0]]
        firstDate = str(row[1][1])[0:10]
        lastDate = str(row[1][2])[0:10]
        geo = str(row[1][3])
        timeframe = firstDate + " " + lastDate
        pytrends.build_payload(kw_list=kw_list,timeframe=timeframe, geo=geo)
        data = pytrends.interest_over_time()
        data = data.drop(labels=['isPartial'],axis='columns')
        data.to_csv(path_or_buf="../downloads/" + kw_list[0] + "_" + firstDate + "-" + lastDate +".csv")

if __name__ == "__main__":
    main()

