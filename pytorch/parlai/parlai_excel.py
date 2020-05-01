"""
Facebook AI powered chatbot in Excel.

There are two functions:
  - parlai_create_world creates a world containing two agents (the AI and the human).
  - parlai_speak takes an input from the human and runs the model to get the AI response.

The entire conversation is returned by parlai_speak so it can be viewed in Excel
rather than just the last response.

See https://parl.ai/projects/blender/
"""
from pyxll import xl_func
from parlai.core.params import ParlaiParser
from parlai.core.agents import create_agent
from parlai.core.worlds import create_task
from excel_agent import HumanExcelAgent


@xl_func("str: object")
def parlai_create_world(model="zoo:blender/blender_90M/model"):
    parser = ParlaiParser(True, True, 'Interactive chat with a model')
    parser.add_argument(
        '-it',
        '--interactive-task',
        type='bool',
        default=True,
        help='Create interactive version of task',
    )
    parser.set_defaults(interactive_mode=True, task='interactive')
    args = ['-t', 'blended_skill_talk', '-mf', model]
    opt = parser.parse_args(print_args=False, args=args)

    agent = create_agent(opt, requireModelExists=True)
    human_agent = HumanExcelAgent(opt)
    world = create_task(opt, [human_agent, agent])
    return world


@xl_func("object, str, int: str[][]")
def parlai_speak(world, input, limit=None):
    human, bot = world.get_agents()[-2:]
    human.set_input(input)
    world.parley()

    messages = [[x.get('id', ''), x.get('text', '')] for x in human.get_conversation()]

    if limit:
        messages = messages[-limit:]
        if len(messages) < limit:
            messages = ([['', '']] * (limit - len(messages))) + messages

    return messages


if __name__ == "__main__":
    world = parlai_create_world()
    conversation = parlai_speak(world, "Hello")
    for id, msg in conversation:
        print(f"{id}: {msg}")
