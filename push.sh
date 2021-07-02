#!/bin/bash

# charliewprice/Guestbook
# clones the GitHub Repository into git/ folder in home directory
# copies files, performs migrations, copies static files
# restarts gunicorn and nginx
#
# (c) 2018, 2019, 2020, 2021 Charlie Price
#

git add .
git commit -m "update"
git push https://charliewprice@github.com/charliewprice/guestbook_analytics.git main
