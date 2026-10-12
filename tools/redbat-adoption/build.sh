#!/bin/bash
# rebuild the harness page into $A/build (the static server reads files per request)
cd /home/dave/Dev/redbat-office-viewers/mega/frontend && npx vite build --config vite.harness.config.ts --outDir /tmp/redbat-adoption/build > /tmp/redbat-adoption/build.log 2>&1; echo "EXIT $?" >> /tmp/redbat-adoption/build.log; tail -2 /tmp/redbat-adoption/build.log
