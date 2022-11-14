#!/bin/bash

set -euo pipefail

dub test --compiler=dmd						   # fast semantic check
dub test -b unittest --compiler=ldc2		   # ASan
dub test -b benchmark-profileGC --compiler=dmd # GC usage
dub test -b benchmark-release --compiler=ldc2  # performance
