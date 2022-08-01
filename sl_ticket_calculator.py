import calendar
import win32com.client
import os
from pathlib import Path

home_path = Path(os.getenv("userprofile"))
txt_file_location = Path(home_path) / Path("LocalDocuments") / Path("EmailScanner") / "Sl_Tickets_Cost.txt"

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")
lasse_katten_account = mapi.Folders("lasse_katten@hotmail.com")
inbox_lasse_katten = lasse_katten_account.Folders("SL Kvitton")
messages = inbox_lasse_katten.Items

total_sum_this_month = 0
total_sum_all_time = 0
current_month_name = ""
index = 0

# Clean out the file
with open(txt_file_location,'w') as f:
    pass

for msg in messages:
  # Dealing with summing up the last messages ticket costs
  if index == len(messages) - 1:
    with open(txt_file_location, "a") as a_file:
      a_file.write(f"\nTicket price sum of {current_month_name} is: {total_sum_this_month}")
      total_sum_all_time += total_sum_this_month
   
  # If we encounter a new month we want to print our results before moving on
  if not calendar.month_name[msg.ReceivedTime.month] == current_month_name:
    with open(txt_file_location, "a") as a_file:
      if current_month_name == "":
        a_file.write(f"Here is the total sum in Kr of all the tickets, sorted by month\n")
      else:
        a_file.write(f"\nTicket price sum of {current_month_name} is: {total_sum_this_month}")
        total_sum_all_time += total_sum_this_month
        total_sum_this_month = 0

  # Getting the sum for one ticket out of the message
  test_split = msg.Body.splitlines()
  for line in test_split:
    if("Att betala" in line):
      stripped_sum_with_kr = line.strip("Att betala").strip()
      complete_stripped_sum = stripped_sum_with_kr.strip("kr").strip()
      print("--- SUM in float : ", float(complete_stripped_sum))
      total_sum_this_month += float(complete_stripped_sum)

      break
  current_month_name = calendar.month_name[msg.ReceivedTime.month]
  index += 1

with open(txt_file_location, "a") as a_file:
  a_file.write(f"\nTotal sum all time: {total_sum_all_time}")

