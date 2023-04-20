from aiogram.dispatcher.filters.state import State, StatesGroup


class DownloadExel(StatesGroup):
    download_exel = State()


class FindValues(StatesGroup):
    find_values = State()
