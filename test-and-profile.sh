#!/bin/bash

set -ex

dub test --compiler=dmd
dub test -b benchmark-release-profileGC --compiler=ldc2
