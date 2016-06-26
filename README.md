# Windows Service Controller

## Summary

Custom service controller.

## Installation

- Download and install the latest [Python 2](https://www.python.org/downloads/).
- Run the following script in commandline.
  - `pip install -r requirements.txt`
- Double-click the **win_service_controller.py** (or use script: `python win_service_controller.py`).

## How To Use

- Setup config file **srv_pak.cfg**, for instance:

    > PackageDisplayName=ServiceName1,ServiceName2...</br>
    > Project_A=MySQL,MSSQLSERVER</br>
    > Project_B=OracleInstance,OracleListener</br>
    > Project_C=OracleInstance,OracleListener,MSSQLSERVER</br>
- Execute the script.

## Functions

- Reload Config: Immediately reload the configuration file **srv_pak.cfg**
- Reload Status: Refresh all package status.
  - Running
  - Stopped
  - Paused
  - Pending - will start/stop after the other services finish
  - Complex - the service status of this package are not sole
  - Config Error
- Start: Start selected packages.
- Advanced Start: Start selected packages, and stop the others.
- Total Start: Start all packages.
- Total Stop: Stop all packages.
