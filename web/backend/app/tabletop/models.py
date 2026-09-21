"""Rule-agnostic tabletop data. Validation limits are technical, not TRPG rules."""
from __future__ import annotations
from typing import Annotated, Literal
from pydantic import BaseModel, ConfigDict, Field

ID = Annotated[str, Field(pattern=r'^[a-zA-Z0-9_-]{1,64}$')]
Coordinate = Annotated[float, Field(ge=-1e7, le=1e7, allow_inf_nan=False)]
Positive = Annotated[float, Field(gt=0, le=1e7, allow_inf_nan=False)]

class Strict(BaseModel):
    model_config = ConfigDict(extra='forbid')

class Counter(Strict):
    name: str = Field(max_length=40)
    value: str = Field(default='', max_length=80)
    maximum: str = Field(default='', max_length=80)

class Piece(Strict):
    id: ID
    kind: Literal['token', 'note', 'circle', 'rect', 'cone', 'line', 'polygon'] = 'token'
    name: str = Field(default='未命名', max_length=100)
    x: Coordinate = 0
    y: Coordinate = 0
    width: Positive = 2
    height: Positive = 2
    radius: Positive = 5
    angle: float = Field(default=60, gt=0, le=360, allow_inf_nan=False)
    rotation: float = Field(default=0, ge=-3600, le=3600, allow_inf_nan=False)
    points: list[tuple[Coordinate, Coordinate]] = Field(default_factory=list, max_length=100)
    color: str = Field(default='#6dcfc1', pattern=r'^#[0-9a-fA-F]{6}$')
    asset: str = Field(default='', pattern=r'^(?:[a-zA-Z0-9_-]{1,64})?$')
    owners: list[ID] = Field(default_factory=list, max_length=24)
    visibility: Literal['all', 'gm', 'selected'] = 'all'
    viewers: list[ID] = Field(default_factory=list, max_length=24)
    locked: bool = False
    enabled: bool = True
    follow: str = Field(default='', pattern=r'^(?:[a-zA-Z0-9_-]{1,64})?$')
    elevation: str = Field(default='', max_length=80)
    footprint: str = Field(default='', max_length=80)
    note: str = Field(default='', max_length=12000)
    gm_note: str = Field(default='', max_length=12000)
    counters: list[Counter] = Field(default_factory=list, max_length=16)
    statuses: list[Annotated[str, Field(max_length=100)]] = Field(default_factory=list, max_length=24)
    source: str = Field(default='', max_length=2000)
    v: int = Field(default=0, ge=0)

class Scene(Strict):
    id: ID
    name: str = Field(default='空白场景', max_length=100)
    grid: bool = False
    snap: bool = False
    grid_size: Positive = 1
    background: str = Field(default='', pattern=r'^(?:[a-zA-Z0-9_-]{1,64})?$')
    map_width: Positive = 60
    map_height: Positive = 40
    round: int = Field(default=1, ge=0, le=1000000)
    turn: str = Field(default='', pattern=r'^(?:[a-zA-Z0-9_-]{1,64})?$')
    order: list[ID] = Field(default_factory=list, max_length=500)
    pieces: dict[str, Piece] = Field(default_factory=dict, max_length=500)
    v: int = Field(default=0, ge=0)

class Edit(Strict):
    id: ID
    expected: int = Field(ge=0)
    value: Piece | None = None

class Command(Strict):
    id: ID
    kind: Literal['pieces', 'scene', 'scene.add', 'scene.delete', 'scene.switch', 'undo', 'member', 'import', 'ping']
    scene: str = ''
    edits: list[Edit] = Field(default_factory=list, max_length=100)
    expected: int = Field(default=0, ge=0)
    data: dict = Field(default_factory=dict)

class Create(Strict):
    name: str = Field(min_length=1, max_length=80)
    nick: str = Field(min_length=1, max_length=40)
    key: str = Field(max_length=256)

class Join(Strict):
    nick: str = Field(min_length=1, max_length=40)
    invite: str = Field(max_length=256)
