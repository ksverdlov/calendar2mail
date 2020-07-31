# calendar2mail

Sends a text report about events in Exchange calendar for the day.

Uses the [keyring|https://pypi.org/project/keyring/] python library as password storage (tested only with **macOS Keychain** as backend).

## Report body example

```
2020-07-31 report

Total time: 10:00:00

09:45 10:00 15M   Some Event Subject
10:00 10:15 15M   Some Event Subject
10:15 11:00 45M   Some Event Subject
11:30 12:15 45M   Some Event Subject
12:15 15:00 2H45M Some Event Subject
15:00 17:00 2H    Some Event Subject
17:00 18:15 1H15M Some Event Subject
18:15 18:30 15M   Some Event Subject
22:00 23:45 1H45M Some Event Subject
```