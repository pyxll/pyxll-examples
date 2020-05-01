"""
A ParlAI agent that takes input from a human Excel user.
Input is set using 'set_input' before calling world.parle.
"""
from parlai.core.agents import Agent
from parlai.core.message import Message


class HumanExcelAgent(Agent):

    def __init__(self, opt):
        super().__init__(opt)
        self.id = 'localExcelHuman'
        self.episodeDone = False
        self.finished = False
        self.__input = ""
        self.__conversation = []

    def set_input(self, input):
        self.__input = input or ""

    def get_conversation(self):
        return self.__conversation

    def epoch_done(self):
        return self.finished

    def observe(self, msg):
        self.__conversation.append(msg)

    def act(self):
        reply = Message()
        reply['id'] = self.getID()
        reply_text = self.__input.replace('\\n', '\n')
        reply['episode_done'] = False
        reply['text'] = reply_text
        self.__conversation.append(reply)
        return reply

    def episode_done(self):
        return self.episodeDone
