import json
with open('resp.json', 'r') as f:
  images = json.load(f)
for image in images:
  image['vulnerabilities'] = list(filter(lambda x:x.get('status', '') != 'open', image['vulnerabilities']))
with open('output.json', 'w') as f:
  f.write(json.dumps(images))
