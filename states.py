from aiogram.dispatcher.filters.state import StatesGroup, State
from aiogram.dispatcher import FSMContext


class Test(StatesGroup):
    wait = State()
    start = State()
    st_count = State()
    st_temp = State()
    st_now = State()
    st_date = State()
    st_work = State()