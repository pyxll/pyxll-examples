# Facebook AI Blender Chatbot

This example requires the 'ParlAI' package from Facebook AI which can be cloned
from https://github.com/facebookresearch/ParlAI.

Check that ParlAI is working correctly by running the following script

```bash
python -Xutf8 parlai/scripts/safe_interactive.py -t blended_skill_talk -mf zoo:blender/blender_90M/model
```

In the example `parlai_excel.py` There are two functions:

- parlai_create_world creates a world containing two agents (the AI and the human).
- parlai_speak takes an input from the human and runs the model to get the AI response.

The entire conversation is returned by parlai_speak so it can be viewed in Excel
rather than just the last response.

## References

https://parl.ai/projects/blender/
https://github.com/facebookresearch/ParlAI
