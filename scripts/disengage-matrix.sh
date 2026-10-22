#!/usr/bin/env bash
# F22 science: map the boundary of the "replace existing value" Disengage trigger.
# Hits the live proxy (real agent+framing path) with a matrix of phrasings/contents,
# classifies each as REFUSE / TOOL_CALL / PROSE, n=REPEAT each, spaced for thread-rate.
#
# Cells encode: label | tools(1/0) | prompt   (tools=0 => no tools => agent-less plain chat)
set -u
BASE="${BASE:-http://localhost:4141/v1}"
REPEAT="${REPEAT:-2}"
COOLDOWN="${COOLDOWN:-12}"
CSV="${CSV:-/tmp/m365-disengage-matrix.csv}"
echo "label,rep,outcome" > "$CSV"

ask() { # $1=tools(1/0) $2=prompt -> echoes outcome
  local tools='[]'
  [ "$1" = 1 ] && tools='[{"type":"function","function":{"name":"bash","description":"Run a shell command","parameters":{"type":"object","properties":{"command":{"type":"string"}},"required":["command"]}}}]'
  # prompts are author-controlled with no " or \ chars, so direct embedding is safe JSON
  local body
  body=$(printf '{"model":"m365-copilot","stream":false,"messages":[{"role":"user","content":"%s"}],"tools":%s}' "$2" "$tools")
  local r; r=$(curl -s -m 90 "$BASE/chat/completions" -H 'Content-Type: application/json' -d "$body")
  if   echo "$r" | grep -q '"disengaged"'; then echo "REFUSE_502"
  elif echo "$r" | grep -q 'tool_calls'; then echo "TOOL_CALL"
  elif echo "$r" | grep -qiE "can.t chat about this|can.t respond to this|different topic|can.t help with"; then echo "REFUSE_apology"
  else echo "PROSE"; fi
}

# label | tools | prompt   (nonce added per-rep to force fresh conversations)
CELLS=(
  "names_old_insteadof|1|Edit config.json so the port is 8080 instead of 3000."
  "names_old_fromto|1|In config.json, change the port from 3000 to 8080."
  "names_old_replacewith|1|In config.json, replace 3000 with 8080."
  "target_only_set|1|Set the port in config.json to 8080."
  "target_only_change|1|Change the port in config.json to 8080."
  "replace_no_oldval|1|In config.json, replace the port value with 8080."
  "code_names_old|1|In calc.py, replace 'return a - b' with 'return a + b'."
  "code_findfix|1|There is a bug in calc.py: the add function subtracts instead of adds. Fix it."
  "text_names_old|1|In greeting.txt, replace the word hello with goodbye."
  "generic_num_old|1|In data.txt, replace 42 with 99."
  "notools_insteadof|0|Edit config.json so the port is 8080 instead of 3000."
  "notools_replacewith|0|In config.json, replace 3000 with 8080."
)

for cell in "${CELLS[@]}"; do
  IFS='|' read -r label tools prompt <<< "$cell"
  for r in $(seq 1 "$REPEAT"); do
    out=$(ask "$tools" "$prompt (probe $label $r-$RANDOM)")
    echo "$label,$r,$out" >> "$CSV"
    printf '  %-22s t=%s #%s -> %s\n' "$label" "$tools" "$r" "$out"
    sleep "$COOLDOWN"
  done
done

echo "=== SUMMARY (label: outcomes) ==="
tail -n +2 "$CSV" | awk -F, '{k=$1; c[k","$3]++; seen[k]=1} END{for(l in seen){printf "  %-22s ",l; for(o in c){split(o,a,",");if(a[1]==l)printf "%s=%s ",a[2],c[o]} print ""}}' | sort
