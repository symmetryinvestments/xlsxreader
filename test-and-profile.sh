#!/bin/bash

set -euo pipefail

dub test --compiler=dmd
dub test -b benchmark-release-profileGC --compiler=ldc2
