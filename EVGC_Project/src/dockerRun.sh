#!/usr/bin/bash

# launches the container after you build it

docker run --rm -it -p 5000:5000 --name ourEVGP docker-evgp
