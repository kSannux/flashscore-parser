from __future__ import annotations

import argparse
import json
import re
import shlex
import sys
import time
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlencode, urljoin

import requests
from bs4 import BeautifulSoup

from flashscore_parser.table_io import fill_xlsx_template, read_xlsx_sheet_templates

FEED_BASE_URL_TEMPLATE = "https://{project_id}.flashscore.ninja/{project_id}/x/feed"
ODDS_API_URL = "https://global.ds.lsapp.eu/odds/pq_graphql"
VALID_BOOK_ID = {16, 419}

def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()

def normalize_args(value: str) -> str:
    normalized = normalize_text(value).casefold().replace("_", "-")
    return normalized

def get_odds(tag: BeautifulSoup) -> tuple[float, float]:
    if (tag.get("title")):
        parts = re.split(r"\s+»\s+", tag.get("title"))
        return map(float, parts)
    
    odd = float(tag.select_one("span").get_text(strip=True))
    return (odd, odd)

def extract_round_number(round_name: str) -> str:
    match = re.search(r"\d+", round_name)
    return match.group() if match else round_name

def tab_url(tabs: dict[str, Link], alias: str) -> str | None:
    link = tabs.get(alias, None)
    return link.url if link and link.url else None

def parse_timestamp(value: str | None) -> str:
    if not value:
        return ""
    try:
        return datetime.fromtimestamp(int(value)).strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return value

def get_correct_numbering(rounds: list[list[Match]]):
    for i in range(len(rounds)):
        rounds[i].reverse()
    rounds.reverse() 

    return rounds

def extract_project_id(html: str) -> str:
    match = re.search(
        r"""["']?projectId["']?\s*:\s*["']?(\d+)""",
        html,
        re.S,
    )
    if match is None:
        raise RuntimeError("projectId не найден в HTML-ответе.")
    return match.group(1)

def build_feed_url(project_id: str | int, feed_name: str) -> str:
    base_url = FEED_BASE_URL_TEMPLATE.format(project_id=project_id)
    return f"{base_url}/{feed_name}"

def parse_league_toggle_key(html: str) -> str:
    match = re.search(r'getToggleIcon\("([^"]+)"', html)
    if not match:
        raise RuntimeError("getToggleIcon league key не найден")

    return match.group(1)

INITIAL_FEED_RE = re.compile(
    r"""cjs\.initialFeeds\[['\"](?P<name>[^'\"]+)['\"]\]\s*=\s*\{\s*data:\s*`(?P<data>[^`]*?)`\s*,\s*
    allEventsCount\s*:\s*(?P<all_events_count>\d+)\s*,\s*
    seasonId\s*:\s*(?P<season_id>\d+)\s*,\s*\}""",
    re.S | re.X,
)

def parse_initial_feed(html: str, name: str) -> dict[str, str]:
    feeds = {
        match.group("name"): {
            "data": match.group("data"),
            "allEventsCount": match.group("all_events_count"),
            "seasonId": match.group("season_id"),
        }
        for match in INITIAL_FEED_RE.finditer(html)
    }
    
    data = feeds.get(name)
    if data is None:
        raise RuntimeError(f"initial feed {name!r} не найден.")
    
    return data

def parse_feed_record(record: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    for part in record.split("¬"):
        if "÷" not in part:
            continue
        key, value = part.split("÷", 1)
        fields[key] = value.replace("\\/", "/")
    return fields

def parse_matches_feed(data: str) -> list[list[Match]]:
    rounds: list[list[Match]] = []
    round_indexes: dict[str, int] = {}

    for record in data.split("¬~"):
        fields = parse_feed_record(record)
        if "AA" not in fields:
            continue

        round_name = extract_round_number(fields.get("ER", ""))
        if round_name not in round_indexes:
            round_indexes[round_name] = len(rounds)
            rounds.append([])

        rounds[round_indexes[round_name]].append(
            Match(
                round=round_name,
                event_id=fields.get("AA", ""),
                team1_participant_id=fields.get("JA", ""),
                team2_participant_id=fields.get("JB", ""),
                time=parse_timestamp(fields.get("AD")),
                team1=fields.get("AE", ""),
                team2=fields.get("AF", ""),
                score1=fields.get("AG", None),
                score2=fields.get("AH", None),
            )
        )

    return rounds

def build_odds_api_url(event_id: str, project_id: str) -> str:
    return f"{ODDS_API_URL}?{urlencode({
        '_hash': 'oce',
        'eventId': event_id,
        'projectId': project_id,
        'geoIpCode': 'CH',
        'geoIpSubdivisionCode': 'CHZH',
    })}"

def odds_pair(item: dict[str, Any]) -> tuple[float, float]:
    value = item.get("value")
    opening = item.get("opening") or value
    try:
        return (float(opening), float(value))
    except (TypeError, ValueError):
        return (0.0, 0.0)

def market_odds(payload: dict[str, Any]) -> list[dict[str, Any]]:
    odds = (
        payload.get("data", {})
        .get("findOddsByEventId", {})
        .get("odds", [])
    )
    if not isinstance(odds, list):
        return []
    return [market for market in odds if isinstance(market, dict)]

def find_market(
    payload: dict[str, Any],
    betting_type: str,
    betting_scope: str = "FULL_TIME",
    book_id: set[int] = VALID_BOOK_ID
) -> dict[str, Any] | None:
    for market in market_odds(payload):
        if (
            market.get("bookmakerId") in book_id
            and market.get("bettingType") == betting_type
            and market.get("bettingScope") == betting_scope
        ):
            return market
    return None

def market_items(market: dict[str, Any] | None) -> list[dict[str, Any]]:
    if not market:
        return []
    odds = market.get("odds", [])
    if not isinstance(odds, list):
        return []
    return [item for item in odds if isinstance(item, dict)]

def handicap_key(item: dict[str, Any]) -> str:
    handicap = item.get("handicap")
    if not isinstance(handicap, dict):
        return ""
    value = handicap.get("value")
    return "" if value is None else str(value)

def apply_participant_market(
    items: list[dict[str, Any]],
    match: Match,
    home_setter,
    away_setter,
    draw_setter=None,
) -> None:
    for item in items:
        participant_id = item.get("eventParticipantId")
        pair = odds_pair(item)
        if participant_id == match.team1_participant_id:
            home_setter(pair)
        elif participant_id == match.team2_participant_id:
            away_setter(pair)
        elif participant_id is None and draw_setter is not None:
            draw_setter(pair)

@dataclass(frozen=True)
class Link:
    text: str
    url: str

@dataclass(frozen=True)
class Match:
    round: str
    event_id: str = ""
    team1_participant_id: str = ""
    team2_participant_id: str = ""
    time: str = ""
    team1: str = ""
    team2: str = ""
    score1: str | None = ""
    score2: str | None = ""


@dataclass(frozen=True)
class H2HMatch:
    time: str
    team1: str
    team2: str
    score1: int
    score2: int
    result: str

    def as_placeholder(self) -> dict[str, object]:
        return {
            "time": self.time,
            "team1": self.team1,
            "team2": self.team2,
            "score1": self.score1,
            "score2": self.score2,
            "score": f"{self.score1}:{self.score2}",
            "result": self.result
        }


@dataclass
class MatchInfo:
    number: int = 0
    round: str = ""
    time: str = ""
    team1: str = ""
    team2: str = ""
    score1: int | None = 0
    score2: int | None = 0
    win1: tuple[float, float] = (0.0, 0.0)
    draw: tuple[float, float] = (0.0, 0.0)
    win2: tuple[float, float] = (0.0, 0.0)
    over: dict[str, tuple[float, float]] = field(default_factory=dict)
    under: dict[str, tuple[float, float]] = field(default_factory=dict)
    both_yes: tuple[float, float] = (0.0, 0.0)
    both_no: tuple[float, float] = (0.0, 0.0)
    double_1x: tuple[float, float] = (0.0, 0.0)
    double_12: tuple[float, float] = (0.0, 0.0)
    double_x2: tuple[float, float] = (0.0, 0.0)
    asian_1: dict[str, tuple[float, float]] = field(default_factory=dict)
    asian_2: dict[str, tuple[float, float]] = field(default_factory=dict)
    european_1: dict[str, tuple[float, float]] = field(default_factory=dict)
    european_x: dict[str, tuple[float, float]] = field(default_factory=dict)
    european_2: dict[str, tuple[float, float]] = field(default_factory=dict)
    no_bet_1: tuple[float, float] = (0.0, 0.0)
    no_bet_2: tuple[float, float] = (0.0, 0.0)
    correct: dict[str, tuple[float, float]] = field(default_factory=dict)
    ht_ft: dict[str, tuple[float, float]] = field(default_factory=dict)
    odd: tuple[float, float] = (0.0, 0.0)
    even: tuple[float, float] = (0.0, 0.0)
    h2h: dict[str, list[H2HMatch]] = field(default_factory=dict)

    def as_placeholder_row(self) -> dict[str, object]:
        return {
            "number": self.number,
            "round": self.round,
            "time": self.time,
            "team1": self.team1,
            "team2": self.team2,
            "score1": self.score1,
            "score2": self.score2,
            "score": f"{self.score1}:{self.score2}",
            "win1": self.win1,
            "win2": self.win2,
            "draw": self.draw,
            "favorite": self.win1 if self.win1 <= self.win2 else self.win2,
            "outsider": self.win1 if self.win1 > self.win2 else self.win2,
            "over": self.over,
            "under": self.under,
            "both-yes": self.both_yes,
            "both-no": self.both_no,
            "double-1x": self.double_1x,
            "double-12": self.double_12,
            "double-x2": self.double_x2,
            "asian-1": self.asian_1,
            "asian-2": self.asian_2,
            "favorite-asian": self.asian_1 if self.win1 > self.win2 else self.asian_2,
            "outsider-asian": self.asian_2 if self.win1 > self.win2 else self.asian_1,
            "european-1": self.european_1,
            "european-x": self.european_x,
            "european-2": self.european_2,
            "no-bet-1": self.no_bet_1,
            "no-bet-2": self.no_bet_2,
            "correct": self.correct,
            "ht-ft": self.ht_ft,
            "odd": self.odd,
            "even": self.even,
            "home": self.h2h["home"],
            "away": self.h2h["away"],
            "h2h": self.h2h["h2h"] 
        }


class FlashscoreClient:
    RETRY_STATUS_CODES = {408, 425, 429, 500, 502, 503, 504}
    BASE_URLS = {"it": "https://www.flashscore.it/",
                 "kz": "https://www.flashscorekz.com/"}

    def __init__(
        self,
        lang: str = "it",
        timeout: float = 20.0,
        max_retries: int = 3,
        retry_delay: float = 1.0,
    ):
        self.base_url = self.BASE_URLS[lang].rstrip("/")
        self.timeout = timeout
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
                ),
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.9",
            }
        )
        self.project_id = extract_project_id(self.fetch())

    def get(self, url: str, headers: dict[str, str] | None = None) -> requests.Response:
        last_exception: requests.RequestException | None = None
        for attempt in range(self.max_retries + 1):
            try:
                response = self.session.get(url, headers=headers, timeout=self.timeout)
                response.raise_for_status()
                return response
            except requests.HTTPError as exc:
                last_exception = exc
                status_code = exc.response.status_code if exc.response is not None else None
                if status_code not in self.RETRY_STATUS_CODES:
                    raise RuntimeError(f"Flashscore вернул HTTP {status_code or 'unknown'} для {url}") from exc
                error = f"HTTP {status_code}"
            except requests.RequestException as exc:
                last_exception = exc
                error = str(exc)

            if attempt == self.max_retries:
                assert last_exception is not None
                raise RuntimeError(
                    f"Не удалось открыть {url} после {self.max_retries + 1} попыток: {error}"
                ) from last_exception

            delay = self.retry_delay * (2 ** attempt)
            print(
                f"Запрос не удался ({error}). Повтор через {delay:g} с "
                f"({attempt + 1}/{self.max_retries})."
            )
            time.sleep(delay)

    def fetch(self, part_url: str = "") -> str:
        url = urljoin(self.base_url, part_url)
        response = self.get(url)
        return response.text
    
    def fetch_feed(self, url: str) -> str:
        response = self.get(
            url,
            headers={
                "Accept": "text/plain, */*;q=0.9",
                "Origin": self.base_url,
                "Referer": f"{self.base_url}/",
                "x-fsign": "SW9D1eZo",
            },
        )
        return response.text

    def fetch_json(self, url: str) -> dict[str, Any]:
        response = self.get(
            url,
            headers={
                "Accept": "application/json,*/*",
                "Origin": self.base_url,
                "Referer": f"{self.base_url}/",
                "x-fsign": "SW9D1eZo",
            },
        )

        try:
            data = response.json()
        except json.JSONDecodeError as exc:
            raise RuntimeError(f"Flashscore вернул не JSON для {url}") from exc
        if not isinstance(data, dict):
            raise RuntimeError(f"Flashscore вернул JSON неожиданного типа для {url}")
        return data

class OddsParserIT:
    def parse_odds(client: FlashscoreClient, match: Match, info: MatchInfo) -> None:
        payload = client.fetch_json(build_odds_api_url(match.event_id, client.project_id))
    
        apply_participant_market(
            market_items(find_market(payload, "HOME_DRAW_AWAY")),
            match,
            lambda pair: setattr(info, "win1", pair),
            lambda pair: setattr(info, "win2", pair),
            lambda pair: setattr(info, "draw", pair),
        )

        for item in market_items(find_market(payload, "OVER_UNDER")):
            key = handicap_key(item)
            if not key:
                continue
            if item.get("selection") == "OVER":
                info.over[key] = odds_pair(item)
            elif item.get("selection") == "UNDER":
                info.under[key] = odds_pair(item)

        for item in market_items(find_market(payload, "BOTH_TEAMS_TO_SCORE")):
            if item.get("bothTeamsToScore") is True:
                info.both_yes = odds_pair(item)
            elif item.get("bothTeamsToScore") is False:
                info.both_no = odds_pair(item)

        apply_participant_market(
            market_items(find_market(payload, "DOUBLE_CHANCE")),
            match,
            lambda pair: setattr(info, "double_1x", pair),
            lambda pair: setattr(info, "double_x2", pair),
            lambda pair: setattr(info, "double_12", pair),
        )

        for item in market_items(find_market(payload, "ASIAN_HANDICAP")):
            key = handicap_key(item)
            if not key:
                continue
            if item.get("eventParticipantId") == match.team1_participant_id:
                info.asian_1[key] = odds_pair(item)
            elif item.get("eventParticipantId") == match.team2_participant_id:
                info.asian_2[key] = odds_pair(item)

        for item in market_items(find_market(payload, "EUROPEAN_HANDICAP")):
            key = handicap_key(item)
            if not key:
                continue
            if item.get("eventParticipantId") == match.team1_participant_id:
                info.european_1[key] = odds_pair(item)
            elif item.get("eventParticipantId") == match.team2_participant_id:
                info.european_2[key] = odds_pair(item)
            elif item.get("eventParticipantId") is None:
                info.european_x[key] = odds_pair(item)

        apply_participant_market(
            market_items(find_market(payload, "DRAW_NO_BET")),
            match,
            lambda pair: setattr(info, "no_bet_1", pair),
            lambda pair: setattr(info, "no_bet_2", pair),
        )

        for item in market_items(find_market(payload, "CORRECT_SCORE")):
            score = item.get("score")
            if score:
                info.correct[str(score)] = odds_pair(item)

        for item in market_items(find_market(payload, "HALF_FULL_TIME")):
            winner = item.get("winner")
            if winner:
                info.ht_ft[str(winner)] = odds_pair(item)

        for item in market_items(find_market(payload, "ODD_OR_EVEN")):
            if item.get("selection") == "ODD":
                info.odd = odds_pair(item)
            elif item.get("selection") == "EVEN":
                info.even = odds_pair(item)


    def parse_h2h(client: FlashscoreClient, match: Match, info: MatchInfo) -> None:
        h2h_url = build_feed_url(client.project_id, f"df_hh_1_{match.event_id}")
        payload = client.fetch_feed(h2h_url)
        
        sections: dict[str, list[H2HMatch]] = {}
        category_index = -1
        section_index = 0
    
        for record in payload.split("¬~"):
            fields = parse_feed_record(record)
            if "KA" in fields:
                category_index += 1
                section_index = 0
            if "KB" in fields:
                section_index += 1
    
            if "KP" not in fields:
                continue
            
            section_name = ""
            if category_index == 0 and section_index == 3:
                section_name = "h2h"
            elif category_index == 1 and section_index == 1:
                section_name = "home"
            elif category_index == 2 and section_index == 1:
                section_name = "away"
    
            if not section_name:
                continue
    
            sections.setdefault(section_name, []).append(
                H2HMatch(
                    time=parse_timestamp(fields.get("KC")),
                    team1=fields.get("KJ", "").lstrip("*"),
                    team2=fields.get("KK", "").lstrip("*"),
                    score1=int(fields.get("KU", "")),
                    score2=int(fields.get("KT", "")),
                    result=fields.get("WIS", ""),
                )
            )

        info.h2h = sections
        

def create_odds_parser(lang: str):
    if lang == "it":
        return OddsParserIT

def parse_fixtures_rounds(html: str) -> list[list[Match]]:
    feed = parse_initial_feed(html, "fixtures")
    rounds = parse_matches_feed(feed["data"])

    return rounds

def parse_results_rounds(client: FlashscoreClient, html: str) -> list[list[Match]]:
    feed = parse_initial_feed(html, "results")

    rounds: list[list[Match]] = []
    league_key = parse_league_toggle_key(html)
    seasod_id = feed.get("seasonId", "")
    project_id = extract_project_id(html)

    i = 0
    response = "none"
    while response:
        api_url = build_feed_url(project_id, f"tr_{league_key}_{seasod_id}_{i}_5_it_1")
        response = client.fetch_feed(api_url)
        
        rounds += parse_matches_feed(response)
        i += 1

    return get_correct_numbering(rounds) 

def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Консольный парсер чемпионатов Flashscore с выгрузкой в xlsx по шаблону."
    )
    parser.add_argument("--template", required=True, help="Путь к шаблону")
    parser.add_argument("--output", default="default", help="Путь к файлу с данными.")
    parser.add_argument("--delay", type=float, default=0, help="Задержка между запросами")
    parser.add_argument("--timeout", type=float, default=50.0, help="Таймаут HTTP-запроса в секундах.")
    return parser


def build_sheet_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--sport", default="football")
    parser.add_argument("--country", required=True)
    parser.add_argument("--league", required=True)
    parser.add_argument("--lang", choices=("it", "kz"), default="it")
    parser.add_argument("--mode", choices=("results", "fixtures", "full"), default="full")
    parser.add_argument("--sheet", required=True)
    return parser


def parse_sheet_args(text: str, source_sheet_name: str) -> argparse.Namespace:
    try:
        return build_sheet_arg_parser().parse_args(shlex.split(text))
    except SystemExit as error:
        raise RuntimeError(f"Некорректные аргументы в листе '{source_sheet_name}': {text}") from error


def collect_matches(sheet_args: argparse.Namespace, delay: float, timeout: float) -> list[MatchInfo]:
    sport = normalize_args(sheet_args.sport)
    country = normalize_args(sheet_args.country)
    league = normalize_args(sheet_args.league)
    client = FlashscoreClient(lang=sheet_args.lang, timeout=timeout)
    odds_parser = create_odds_parser(sheet_args.lang)

    url = f"/{sport}/{country}/{league}"
    league_html = client.fetch(url)

    result_rounds = parse_results_rounds(client, league_html)
    fixture_rounds = parse_fixtures_rounds(league_html)
    total_matches = sum(len(round_matches) for round_matches in result_rounds)
    if not result_rounds or total_matches == 0:
        raise RuntimeError(
            f"На странице результатов не найдено матчей: {url}/results. "
            "Проверьте предыдущие сезоны в archive и укажите в --league"
        )

    matches_info: list[MatchInfo] = []
    cnt = 1

    if sheet_args.mode in {"results", "full"}:
        for round_matches in result_rounds:
            for match in round_matches:
                info = MatchInfo(
                    number=cnt,
                    round=match.round,
                    time=match.time,
                    team1=match.team1,
                    team2=match.team2,
                    score1=int(match.score1),
                    score2=int(match.score2),
                )
                cnt += 1

                odds_parser.parse_odds(client, match, info)
                if delay > 0:
                    time.sleep(delay)

                odds_parser.parse_h2h(client, match, info)

                matches_info.append(info)
                print(
                    f"Обработан матч {info.number}: "
                    f"тур {info.round}, {info.time}, "
                    f"{info.team1} {info.score1}:{info.score2} {info.team2}"
                )

    if sheet_args.mode in {"fixtures", "full"}:
        for round_matches in fixture_rounds:
            for match in round_matches:
                info = MatchInfo(
                    number=cnt,
                    round=match.round,
                    time=match.time,
                    team1=match.team1,
                    team2=match.team2,
                    score1=match.score1,
                    score2=match.score2,
                )
                cnt += 1

                odds_parser.parse_odds(client, match, info)
                if not all(pair[0] > 0 and pair[1] > 0 for pair in (info.win1, info.draw, info.win2)):
                    return matches_info
                if delay > 0:
                    time.sleep(delay)

                odds_parser.parse_h2h(client, match, info)

                matches_info.append(info)
                print(
                    f"Обработан матч {info.number}: "
                    f"тур {info.round}, {info.time}, "
                    f"{info.team1} {info.score1}:{info.score2} {info.team2}"
                )

    return matches_info


def main(argv: list[str] | None = None) -> int:
    args = build_arg_parser().parse_args(argv)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    template_suffix = Path(args.template).suffix.lower()
    output = f"flashscore-{timestamp}{template_suffix}" if args.output == "default" else args.output

    try:
        sheet_templates = read_xlsx_sheet_templates(args.template)
        sheet_jobs = [
            (template, parse_sheet_args(config_text, template.sheet_name))
            for template, config_text in sheet_templates
        ]
        target_names = [sheet_args.sheet for _, sheet_args in sheet_jobs]
        if len(target_names) != len(set(target_names)):
            raise RuntimeError("Значения --sheet должны быть уникальными для всех листов.")
        source_names = {template.sheet_name for template, _ in sheet_jobs}
        for template, sheet_args in sheet_jobs:
            if sheet_args.sheet in source_names and sheet_args.sheet != template.sheet_name:
                raise RuntimeError(
                    f"Лист '{template.sheet_name}' нельзя переименовать в '{sheet_args.sheet}': "
                    "это имя другого исходного листа."
                )

        for index, (template, sheet_args) in enumerate(sheet_jobs):
            print(f"Лист {template.sheet_name}: сбор данных для {sheet_args.sport}/{sheet_args.country}/{sheet_args.league}.")
            matches_info = collect_matches(sheet_args, args.delay, args.timeout)
            print(f"Лист {template.sheet_name}: данные собраны. Заполнение файла...", flush=True)
            fill_xlsx_template(
                [info.as_placeholder_row() for info in matches_info],
                template,
                args.template,
                output,
                copy_template=index == 0,
                target_sheet_name=sheet_args.sheet,
            )
            print(f"Лист {sheet_args.sheet}: сохранено {len(matches_info)} записей.", flush=True)
    except RuntimeError as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1

    print(f"Готово: все листы сохранены в {output}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
