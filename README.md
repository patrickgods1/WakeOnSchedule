# WakeOnSchedule
WakeOnSchedule is an application designed to automate the following:

* Send a magic packet based on the daily schedule report to wake computers via wake on lan.

## Development
These instructions will get you a copy of the project up and running on your local machine for development.

### Built With
* [Python 3.8](https://docs.python.org/3/) - The scripting language used.
* [Pandas](https://pandas.pydata.org/) - Data structure/anaylsis tool used.
* [xlrd](https://xlrd.readthedocs.io/en/latest/index.html) - Used to read Microsoft Excel documents (Daily Schedule)

### Running the Script
Run the following command to installer all the required Python modules:
```
pip install -r requirements.txt
```
Modify the config.py to specify the:
* BROADCAST_IP -> str
* DEFAULT_PORT -> int
* hwAddress -> Dict(str, list[str])

To run the application:
```
python .\WakeOnSchedule.py
```

## Authors
* **Patrick Yu** - *Initial work* - [patrickgods1](https://github.com/patrickgods1) - UC Berkeley Extension
