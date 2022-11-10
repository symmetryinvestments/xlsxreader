#!/bin/bash

set -euo pipefail

dub test --compiler=dmd
dub test -b benchmark-profileGC --compiler=dmd
dub test -b benchmark-release --compiler=ldc2
