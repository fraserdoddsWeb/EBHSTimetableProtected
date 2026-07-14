from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Tuple, Optional
from collections import defaultdict

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
WEEKS = ["A", "B"]


@dataclass
class SolveResult:
    status: str
    message: str
    allocations: List[Dict[str, Any]]
    teacher_daily: List[Dict[str, Any]]
    class_summary: List[Dict[str, Any]]
    diagnostics: List[str]
    objective: Optional[float] = None


def _clean(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _truthy(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return _clean(value).lower() in {"yes", "true", "1", "y"}


def _to_int(value: Any, default: int = 0) -> int:
    try:
        if value is None or value == "":
            return default
        return int(float(value))
    except Exception:
        return default


def _priority_weight(value: Any) -> int:
    v = _clean(value).lower()
    if v in {"very high", "very_high", "highest"}:
        return 1200
    if v in {"high"}:
        return 650
    return 300


def _slot_key(row: Dict[str, Any]) -> Tuple[str, str, str]:
    return (_clean(row.get("Week")), _clean(row.get("Day")), _clean(row.get("Period")))


def _subjects_compatible(teacher_subject: str, class_subject: str) -> bool:
    """Strict subject compatibility for teacher allocation.

    Earlier versions used broad pools, which was useful for producing a quick
    diagnostic export but created unrealistic splits such as Food Tech being
    allocated to Art or Science staff. Broad pools are still used by the timing
    builder for load estimates, but actual teacher allocation now uses this
    stricter map unless the user explicitly allows emergency non-specialists.
    """
    ts = _clean(teacher_subject).lower().replace("&", "and")
    cs = _clean(class_subject).lower().replace("&", "and")
    if cs == "pshe":
        return True
    if not ts or ts in {"any", "all"}:
        return True
    if ts == cs:
        return True

    # KS3 Science staff can teach generic Science. KS4 strands are specialism
    # based, so Chemistry is not automatically Biology, etc.
    if cs == "science" and ts in {"science", "biology", "chemistry", "physics"}:
        return True
    if cs in {"biology", "chemistry", "physics"}:
        return ts == cs or ts == "science specialist"

    pe = {"p.e", "pe", "p e", "physical education", "p.e - gcse", "pe - btec"}
    if ts in pe and cs in pe:
        return True
    art = {"art", "art + design", "art and design"}
    if ts in art and cs in art:
        return True
    computing = {"computing", "computer science", "creative imedia"}
    if ts in computing and cs in computing:
        return True
    languages = {"french", "german", "spanish"}
    if ts in languages and cs in languages:
        return True
    food = {"food tech", "hospitality"}
    if ts in food and cs in food:
        return True
    # Business and Technology stay separate unless a teacher explicitly has them.
    return False


def _teacher_can_teach(teacher_row: Dict[str, Any], class_subject: str) -> bool:
    subjects = _clean(teacher_row.get("Subjects"))
    if not subjects or subjects.lower() in {"any", "all"}:
        return True
    return any(_subjects_compatible(s, class_subject) for s in subjects.replace(";", ",").split(",") if _clean(s))


def _build_maps(project: Dict[str, Any]) -> Dict[str, Any]:
    teachers = [dict(r) for r in project.get("teachers", []) if _clean(r.get("Teacher"))]
    classes = [dict(r) for r in project.get("classes", []) if _clean(r.get("ClassID"))]
    lessons = [dict(r) for r in project.get("lessons", []) if _clean(r.get("ClassID"))]
    periods = [dict(r) for r in project.get("periods", []) if _clean(r.get("Period"))]
    fixed = [dict(r) for r in project.get("non_teaching", []) if _clean(r.get("Teacher"))]
    rules = [dict(r) for r in project.get("teacher_rules", []) if _clean(r.get("ClassID")) and _clean(r.get("Teacher"))]
    roles = [dict(r) for r in project.get("teacher_roles", []) if _clean(r.get("Teacher"))]

    teacher_map = {_clean(t.get("Teacher")): t for t in teachers}
    class_map = {_clean(c.get("ClassID")): c for c in classes}
    period_order = {_clean(p.get("Period")): i for i, p in enumerate(periods)}
    teaching_periods = [_clean(p.get("Period")) for p in periods if _truthy(p.get("Teaching"))]

    return {
        "teachers": teachers,
        "classes": classes,
        "lessons": lessons,
        "periods": periods,
        "fixed": fixed,
        "rules": rules,
        "roles": roles,
        "teacher_role_map": {_clean(r.get("Teacher")): r for r in roles},
        "teacher_map": teacher_map,
        "class_map": class_map,
        "period_order": period_order,
        "teaching_periods": teaching_periods,
    }


def solve_project(project: Dict[str, Any], time_limit_seconds: int = 30, max_teacher_attempts: Optional[List[int]] = None) -> SolveResult:
    """Solve a fixed lesson teacher allocation problem.

    This version assumes the lesson slots are fixed. It optimises teacher allocation.
    It tries to keep every class within the requested maximum teacher count.
    If that fails, it can relax the maximum teacher count to 3.
    """
    try:
        from ortools.sat.python import cp_model
    except Exception as exc:
        return SolveResult(
            status="Missing dependency",
            message="OR Tools is not installed. Run: pip install -r requirements.txt",
            allocations=[],
            teacher_daily=[],
            class_summary=[],
            diagnostics=[str(exc)],
        )

    maps = _build_maps(project)
    teachers = maps["teachers"]
    classes = maps["classes"]
    lessons = maps["lessons"]
    fixed = maps["fixed"]
    rules = maps["rules"]
    teacher_role_map = maps.get("teacher_role_map", {})
    teacher_map = maps["teacher_map"]
    class_map = maps["class_map"]
    period_order = maps["period_order"]
    teaching_periods = maps["teaching_periods"]

    teacher_names = list(teacher_map.keys())
    if not teacher_names:
        return SolveResult("No teachers", "Add at least one teacher.", [], [], [], ["No teachers were entered."])
    if not lessons:
        return SolveResult("No lessons", "Add at least one fixed lesson row.", [], [], [], ["No lessons were entered."])

    task_rows = []
    diagnostics = []
    for i, row in enumerate(lessons):
        class_id = _clean(row.get("ClassID"))
        if class_id not in class_map:
            diagnostics.append(f"Lesson row {i + 1} uses unknown class {class_id}.")
            continue
        period = _clean(row.get("Period"))
        if period not in teaching_periods:
            diagnostics.append(f"Lesson row {i + 1} is ignored because {period} is not a teaching period.")
            continue
        task_rows.append({**row, "TaskID": len(task_rows)})

    if not task_rows:
        return SolveResult("No valid lessons", "No valid teaching lesson rows were found.", [], [], [], diagnostics)

    unavailable = set()
    fixed_by_teacher_slot = {}
    for row in fixed:
        teacher = _clean(row.get("Teacher"))
        slot = _slot_key(row)
        unavailable.add((teacher, slot))
        fixed_by_teacher_slot[(teacher, slot)] = _clean(row.get("Reason")) or "Fixed"

    rules_by_class = defaultdict(list)
    for rule in rules:
        rules_by_class[_clean(rule.get("ClassID"))].append(rule)

    if max_teacher_attempts is None:
        # Try progressively. Many projects use MaxTeachers = 1 to mean
        # "ideal one teacher", not a hard impossibility. So with relaxation on,
        # try one teacher first, then two, then three.
        requested = []
        for c in classes:
            requested.append(_to_int(c.get("MaxTeachers"), 1) or 1)
        base = max(1, min(requested) if requested else 1)
        max_teacher_attempts = []
        for cap in [base, 2, 3]:
            if cap not in max_teacher_attempts:
                max_teacher_attempts.append(cap)

    last_message = ""
    for max_teacher_cap in max_teacher_attempts:
        result = _solve_once(
            cp_model,
            project,
            task_rows,
            teachers,
            classes,
            teacher_names,
            teacher_map,
            teacher_role_map,
            class_map,
            rules_by_class,
            unavailable,
            fixed_by_teacher_slot,
            teaching_periods,
            period_order,
            max_teacher_cap,
            time_limit_seconds,
        )
        if result.status in {"Optimal", "Feasible"}:
            if max_teacher_cap > 2:
                result.diagnostics.insert(0, f"The solver had to relax the class teacher maximum to {max_teacher_cap}.")
            result.diagnostics.extend(diagnostics)
            return result
        last_message = result.message

    # If the strict global model cannot prove a solution, produce a diagnostic
    # working export using slot-by-slot matching rather than failing completely.
    fallback_cap = max(max_teacher_attempts or [3])
    fallback_result = _fallback_slot_match(project, max_teacher_cap=fallback_cap, load_slack=5)
    if fallback_result.status in {"Optimal", "Feasible"}:
        fallback_result.diagnostics = diagnostics + [last_message] + fallback_result.diagnostics
        return fallback_result

    return SolveResult(
        status="Infeasible",
        message=last_message or "No valid allocation found. Try relaxing teacher limits, fixed allocations or class maximum teacher counts.",
        allocations=[],
        teacher_daily=[],
        class_summary=[],
        diagnostics=diagnostics + [last_message] + fallback_result.diagnostics,
    )



def _fallback_slot_match(project: Dict[str, Any], max_teacher_cap: int = 3, load_slack: int = 5, time_limit_per_slot: int = 3) -> SolveResult:
    """Last-resort timetable assignment fallback.

    It solves each populated period as a small matching problem. Subject specialists
    are always preferred, and non-subject emergency assignments are heavily
    penalised and reported in the diagnostics rather than hidden. This gives a
    usable export for diagnosis when the full CP-SAT allocation model is too tight.
    """
    try:
        from ortools.sat.python import cp_model
    except Exception as exc:
        return SolveResult("Missing dependency", "OR Tools is not installed.", [], [], [], [str(exc)])

    maps = _build_maps(project)
    teachers = maps["teachers"]
    teacher_map = maps["teacher_map"]
    class_map = maps["class_map"]
    teaching_periods = maps["teaching_periods"] or ["P1", "P2", "P3", "P4", "P5"]
    lessons = maps["lessons"]
    fixed = maps["fixed"]
    unavailable = {(_clean(r.get("Teacher")), _slot_key(r)) for r in fixed}

    lessons_by_slot: Dict[Tuple[str, str, str], List[Dict[str, Any]]] = defaultdict(list)
    for lesson in lessons:
        if _clean(lesson.get("ClassID")) in class_map:
            lessons_by_slot[_slot_key(lesson)].append(dict(lesson))

    period_order = {p: i for i, p in enumerate(teaching_periods)}
    slots = sorted(lessons_by_slot, key=lambda s: (s[0], DAYS.index(s[1]) if s[1] in DAYS else 99, period_order.get(s[2], 99)))

    load = defaultdict(int)
    target = {name: _to_int(row.get("TargetLessons"), 0) for name, row in teacher_map.items()}
    max_load = {name: max(_to_int(row.get("MaxLessons"), target[name]), target[name]) + load_slack for name, row in teacher_map.items()}
    class_teachers: Dict[str, set] = defaultdict(set)
    allocations: List[Dict[str, Any]] = []
    emergency_count = 0
    over_load_count = 0

    for slot in slots:
        slot_lessons = lessons_by_slot[slot]
        model = cp_model.CpModel()
        x = {}
        candidates: Dict[int, List[str]] = {}
        emergency_flag: Dict[Tuple[int, str], int] = {}

        for i, lesson in enumerate(slot_lessons):
            class_id = _clean(lesson.get("ClassID"))
            subject = _clean(class_map[class_id].get("Subject"))
            specialist_candidates = []
            emergency_candidates = []
            allow_emergency = bool(project.get("settings", {}).get("fallback_allow_emergency_non_subject", False))
            class_year = _to_int(class_map[class_id].get("Year"), 0)
            for teacher, trow in teacher_map.items():
                if (teacher, slot) in unavailable:
                    continue
                if _teacher_can_teach(trow, subject):
                    specialist_candidates.append(teacher)
                elif allow_emergency and class_year in {7, 8, 9}:
                    # Emergency non-specialist cover is never allowed for KS4.
                    # If the user allows it, use it only as a last resort in KS3.
                    # The objective below makes Year 7 cheapest, then Year 8, then Year 9.
                    emergency_candidates.append(teacher)
            all_candidates = specialist_candidates + emergency_candidates
            if not all_candidates:
                return SolveResult("Infeasible", f"Fallback could not find any specialist teacher for {class_id} ({subject}) in {slot}.", [], [], [], [])
            candidates[i] = all_candidates
            for teacher in all_candidates:
                x[(i, teacher)] = model.NewBoolVar(f"fallback_{i}_{abs(hash(teacher)) % 1000000}")
                emergency_flag[(i, teacher)] = 0 if teacher in specialist_candidates else 1
            model.Add(sum(x[(i, teacher)] for teacher in all_candidates) == 1)

        for teacher in teacher_map:
            vars_here = [x[(i, teacher)] for i in range(len(slot_lessons)) if (i, teacher) in x]
            if vars_here:
                model.Add(sum(vars_here) <= 1)

        objective_terms = []
        for i, lesson in enumerate(slot_lessons):
            class_id = _clean(lesson.get("ClassID"))
            current = _clean(lesson.get("CurrentTeacher"))
            used_for_class = class_teachers[class_id]
            for teacher in candidates[i]:
                penalty = 0
                if emergency_flag[(i, teacher)]:
                    class_year = _to_int(class_map[class_id].get("Year"), 0)
                    if class_year >= 10:
                        penalty += 10000000
                    else:
                        # Emergency cover priority: Year 7, then Year 8, then Year 9.
                        penalty += 120000 + max(0, class_year - 7) * 120000
                if teacher in used_for_class:
                    penalty -= 10000
                elif len(used_for_class) >= max_teacher_cap:
                    penalty += 30000
                else:
                    penalty += 1000 * len(used_for_class)
                if current and teacher == current:
                    penalty -= 500
                teacher_target = max(1, target.get(teacher, 0) or max_load.get(teacher, 1))
                projected_load = load[teacher] + 1
                if projected_load > max_load[teacher]:
                    penalty += 80000 * (projected_load - max_load[teacher])
                # Normalised load balancing. This stops small-allocation staff
                # being used too quickly just because their raw load is low.
                penalty += int(2500 * projected_load * projected_load / teacher_target)
                if load[teacher] < teacher_target:
                    penalty -= 25 * (teacher_target - load[teacher])
                objective_terms.append(penalty * x[(i, teacher)])
        model.Minimize(sum(objective_terms))

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit_per_slot
        solver.parameters.num_search_workers = 8
        status = solver.Solve(model)
        if status not in {cp_model.OPTIMAL, cp_model.FEASIBLE}:
            return SolveResult("Infeasible", f"Fallback could not match all lessons in {slot}.", [], [], [], [])

        for i, lesson in enumerate(slot_lessons):
            chosen = ""
            for teacher in candidates[i]:
                if solver.Value(x[(i, teacher)]) == 1:
                    chosen = teacher
                    break
            class_id = _clean(lesson.get("ClassID"))
            subject = _clean(class_map[class_id].get("Subject"))
            if emergency_flag[(i, chosen)]:
                emergency_count += 1
            if load[chosen] + 1 > max_load[chosen]:
                over_load_count += 1
            load[chosen] += 1
            class_teachers[class_id].add(chosen)
            allocations.append({
                "Week": _clean(lesson.get("Week")),
                "Day": _clean(lesson.get("Day")),
                "Period": _clean(lesson.get("Period")),
                "SlotID": f"{_clean(lesson.get('Week'))}_{_clean(lesson.get('Day'))}_{_clean(lesson.get('Period'))}",
                "Subject": subject,
                "Block": _clean(lesson.get("Block")),
                "ClassID": class_id,
                "CurrentTeacher": _clean(lesson.get("CurrentTeacher")),
                "NewTeacher": chosen,
            })

    teacher_daily = []
    for teacher in teacher_map:
        for week in WEEKS:
            for day in DAYS:
                day_slots = [(week, day, p) for p in teaching_periods]
                available = [s for s in day_slots if (teacher, s) not in unavailable]
                lesson_count = sum(1 for a in allocations if a["NewTeacher"] == teacher and (a["Week"], a["Day"], a["Period"]) in available)
                teacher_daily.append({
                    "Teacher": teacher,
                    "Week": week,
                    "Day": day,
                    "AvailablePeriods": len(available),
                    "Lessons": lesson_count,
                    "Frees": len(available) - lesson_count,
                    "FullDay": "Yes" if available and lesson_count == len(available) else "No",
                    "Role": "",
                    "ProtectedRole": "No",
                    "TargetFreesPerDay": "",
                    "MeetsProtectedFreeTarget": "Yes",
                })

    class_summary = []
    for class_id in sorted(class_teachers):
        current_counts = defaultdict(int)
        new_counts = defaultdict(int)
        for lesson in lessons:
            if _clean(lesson.get("ClassID")) == class_id:
                current = _clean(lesson.get("CurrentTeacher"))
                if current:
                    current_counts[current] += 1
        for allocation in allocations:
            if allocation["ClassID"] == class_id:
                new_counts[allocation["NewTeacher"]] += 1
        def fmt_counts(d: Dict[str, int]) -> str:
            return ", ".join(f"{t} ({n})" for t, n in sorted(d.items(), key=lambda kv: (-kv[1], kv[0])))
        class_summary.append({
            "ClassID": class_id,
            "Block": next((a["Block"] for a in allocations if a["ClassID"] == class_id), ""),
            "Lessons": sum(1 for a in allocations if a["ClassID"] == class_id),
            "CurrentAllocation": fmt_counts(current_counts),
            "NewAllocation": fmt_counts(new_counts),
            "TeacherCount": len(new_counts),
            "Status": "One teacher" if len(new_counts) == 1 else "Split",
        })

    diagnostics = [
        "Used fallback slot matching because the strict CP-SAT teacher allocation was infeasible on the generated timing grid.",
        f"Emergency non-subject assignments used: {emergency_count}.",
        f"Assignments above max/target plus fallback slack: {over_load_count}.",
        "Emergency non-specialist fallback, when enabled, is limited to KS3 and is penalised in this order: Year 7, then Year 8, then Year 9. It is never allowed for Year 10 or Year 11.",
        "Use this export as a working diagnostic timetable. Reduce the emergency count by adding more subject capacity, widening teacher subject permissions, or regenerating timings.",
    ]
    return SolveResult("Feasible", "Solved with fallback slot matching. Review diagnostics before treating this as final.", allocations, teacher_daily, class_summary, diagnostics)

def _solve_once(
    cp_model,
    project: Dict[str, Any],
    task_rows: List[Dict[str, Any]],
    teachers: List[Dict[str, Any]],
    classes: List[Dict[str, Any]],
    teacher_names: List[str],
    teacher_map: Dict[str, Dict[str, Any]],
    teacher_role_map: Dict[str, Dict[str, Any]],
    class_map: Dict[str, Dict[str, Any]],
    rules_by_class: Dict[str, List[Dict[str, Any]]],
    unavailable: set,
    fixed_by_teacher_slot: Dict[Tuple[str, Tuple[str, str, str]], str],
    teaching_periods: List[str],
    period_order: Dict[str, int],
    max_teacher_cap: int,
    time_limit_seconds: int,
) -> SolveResult:
    model = cp_model.CpModel()
    x = {}
    feasible_teachers_by_task = {}
    diagnostics = []

    slot_by_task = {t["TaskID"]: _slot_key(t) for t in task_rows}
    class_by_task = {t["TaskID"]: _clean(t.get("ClassID")) for t in task_rows}

    hard_not_allowed = set()
    allowed_rules_by_class = defaultdict(set)
    exact_rules = []
    preferred_rules = []

    for class_id, class_rules in rules_by_class.items():
        for r in class_rules:
            teacher = _clean(r.get("Teacher"))
            rule_type = _clean(r.get("Rule")).lower()
            exact = _to_int(r.get("ExactLessons"), 0)
            if rule_type == "not allowed":
                hard_not_allowed.add((class_id, teacher))
            elif rule_type == "allowed":
                allowed_rules_by_class[class_id].add(teacher)
            elif rule_type == "must":
                exact_rules.append((class_id, teacher, exact))
            elif rule_type == "preferred":
                preferred_rules.append((class_id, teacher, _to_int(r.get("Penalty"), 30)))

    for task in task_rows:
        tid = task["TaskID"]
        class_id = class_by_task[tid]
        class_row = class_map[class_id]
        class_subject = _clean(class_row.get("Subject"))
        slot = slot_by_task[tid]
        feasible = []
        for teacher in teacher_names:
            trow = teacher_map[teacher]
            if (teacher, slot) in unavailable:
                continue
            if not _teacher_can_teach(trow, class_subject):
                continue
            if (class_id, teacher) in hard_not_allowed:
                continue
            if class_id in allowed_rules_by_class and teacher not in allowed_rules_by_class[class_id]:
                continue
            var = model.NewBoolVar(f"x_t{tid}_{teacher}")
            x[(tid, teacher)] = var
            feasible.append(teacher)
        feasible_teachers_by_task[tid] = feasible
        if not feasible:
            diagnostics.append(f"No feasible teacher for {class_id} in {slot}.")
            return SolveResult("Infeasible", diagnostics[-1], [], [], [], diagnostics)
        model.Add(sum(x[(tid, teacher)] for teacher in feasible) == 1)

    tasks_by_slot = defaultdict(list)
    for task in task_rows:
        tasks_by_slot[slot_by_task[task["TaskID"]]].append(task["TaskID"])

    teach_slot = {}
    all_slots = sorted(set(slot_by_task.values()) | {(w, d, p) for w in WEEKS for d in DAYS for p in teaching_periods}, key=lambda s: (s[0], DAYS.index(s[1]) if s[1] in DAYS else 99, period_order.get(s[2], 99)))
    for teacher in teacher_names:
        for slot in all_slots:
            vars_here = [x[(tid, teacher)] for tid in tasks_by_slot[slot] if (tid, teacher) in x]
            b = model.NewBoolVar(f"teach_{teacher}_{slot[0]}_{slot[1]}_{slot[2]}")
            teach_slot[(teacher, slot)] = b
            if vars_here:
                model.AddMaxEquality(b, vars_here)
                model.Add(sum(vars_here) <= 1)
            else:
                model.Add(b == 0)

    class_teacher_used = {}
    tasks_by_class = defaultdict(list)
    for task in task_rows:
        tasks_by_class[class_by_task[task["TaskID"]]].append(task["TaskID"])

    for class_id, tids in tasks_by_class.items():
        used_vars = []
        for teacher in teacher_names:
            assigns = [x[(tid, teacher)] for tid in tids if (tid, teacher) in x]
            used = model.NewBoolVar(f"used_{class_id}_{teacher}")
            class_teacher_used[(class_id, teacher)] = used
            if assigns:
                model.AddMaxEquality(used, assigns)
            else:
                model.Add(used == 0)
            used_vars.append(used)
        c_row = class_map[class_id]
        requested_max = _to_int(c_row.get("MaxTeachers"), 1) or 1
        # max_teacher_cap is the relaxation level for this solve attempt.
        # If requested_max is 1, the solver first tries 1, then relaxes to 2/3
        # only if needed. If a user has explicitly put a higher cap, keep it.
        effective_max = max(max_teacher_cap, requested_max)
        model.Add(sum(used_vars) <= effective_max)

    for class_id, teacher, exact in exact_rules:
        tids = tasks_by_class.get(class_id, [])
        if not tids:
            continue
        assigns = [x[(tid, teacher)] for tid in tids if (tid, teacher) in x]
        if exact > 0:
            model.Add(sum(assigns) == exact)
        else:
            model.Add(sum(assigns) == len(tids))

    objectives = []

    for class_id, tids in tasks_by_class.items():
        teacher_count = model.NewIntVar(0, len(teacher_names), f"teacher_count_{class_id}")
        model.Add(teacher_count == sum(class_teacher_used[(class_id, teacher)] for teacher in teacher_names))
        year = _to_int(class_map[class_id].get("Year"), 0)
        weight = 1200 if year in {10, 11} else 850
        objectives.append(weight * teacher_count)
        if max_teacher_cap > 2:
            extra = model.NewIntVar(0, len(teacher_names), f"extra_teachers_{class_id}")
            zero_const = model.NewConstant(0)
            diff = model.NewIntVar(-2, len(teacher_names), f"diff_teachers_{class_id}")
            model.Add(diff == teacher_count - 2)
            model.AddMaxEquality(extra, [diff, zero_const])
            objectives.append(5000 * extra)

    for teacher in teacher_names:
        actual = model.NewIntVar(0, len(task_rows), f"load_{teacher}")
        teacher_assigns = [x[(tid, teacher)] for tid in range(len(task_rows)) if (tid, teacher) in x]
        model.Add(actual == sum(teacher_assigns))
        target = _to_int(teacher_map[teacher].get("TargetLessons"), 0)
        max_load = _to_int(teacher_map[teacher].get("MaxLessons"), 0)
        if max_load > 0:
            model.Add(actual <= max_load)
        if target > 0:
            over = model.NewIntVar(0, len(task_rows), f"over_{teacher}")
            under = model.NewIntVar(0, len(task_rows), f"under_{teacher}")
            model.Add(actual - target == over - under)
            objectives.append(250 * over)
            objectives.append(250 * under)

    for task in task_rows:
        tid = task["TaskID"]
        current_teacher = _clean(task.get("CurrentTeacher"))
        if current_teacher and (tid, current_teacher) in x:
            year = _to_int(class_map[class_by_task[tid]].get("Year"), 0)
            penalty = 120 if year in {10, 11} else 15
            objectives.append(penalty * (1 - x[(tid, current_teacher)]))

    for class_id, teacher, penalty in preferred_rules:
        tids = tasks_by_class.get(class_id, [])
        for tid in tids:
            if (tid, teacher) in x:
                objectives.append(penalty * (1 - x[(tid, teacher)]))

    # Teacher welfare score
    days = [(w, d) for w in WEEKS for d in DAYS]
    for teacher in teacher_names:
        for week, day in days:
            day_slots = [(week, day, p) for p in teaching_periods if (week, day, p) in all_slots]
            available_day_slots = [s for s in day_slots if (teacher, s) not in unavailable]
            if not available_day_slots:
                continue
            lessons_day = model.NewIntVar(0, len(available_day_slots), f"lessons_{teacher}_{week}_{day}")
            model.Add(lessons_day == sum(teach_slot[(teacher, s)] for s in available_day_slots))

            excess_four = model.NewIntVar(0, len(available_day_slots), f"excess_four_{teacher}_{week}_{day}")
            diff = model.NewIntVar(-5, 5, f"diff_four_{teacher}_{week}_{day}")
            model.Add(diff == lessons_day - 4)
            model.AddMaxEquality(excess_four, [diff, model.NewConstant(0)])
            objectives.append(90 * excess_four)

            ordered = sorted(available_day_slots, key=lambda s: period_order.get(s[2], 99))
            for start in range(max(0, len(ordered) - 3)):
                window = ordered[start:start+4]
                run = model.NewBoolVar(f"run4_{teacher}_{week}_{day}_{start}")
                model.Add(sum(teach_slot[(teacher, s)] for s in window) == 4).OnlyEnforceIf(run)
                model.Add(sum(teach_slot[(teacher, s)] for s in window) <= 3).OnlyEnforceIf(run.Not())
                objectives.append(70 * run)

            role = teacher_role_map.get(teacher)
            protected = role and _truthy(role.get("Protected", True))
            if protected:
                priority = _priority_weight(role.get("Priority"))
                target_frees = max(0, _to_int(role.get("TargetFreesPerDay"), 1))
                max_consecutive = max(1, _to_int(role.get("MaxConsecutiveLessons"), 3))

                if target_frees > 0 and len(available_day_slots) > target_frees:
                    # Strongly prefer at least target_frees free periods on each available day.
                    max_lessons_for_day = max(0, len(available_day_slots) - target_frees)
                    no_hoy_free_excess = model.NewIntVar(0, len(available_day_slots), f"hoy_free_excess_{teacher}_{week}_{day}")
                    hoy_diff = model.NewIntVar(-len(available_day_slots), len(available_day_slots), f"hoy_free_diff_{teacher}_{week}_{day}")
                    model.Add(hoy_diff == lessons_day - max_lessons_for_day)
                    model.AddMaxEquality(no_hoy_free_excess, [hoy_diff, model.NewConstant(0)])
                    objectives.append(priority * no_hoy_free_excess)

                # Strongly avoid long consecutive runs for protected roles.
                run_length = max_consecutive + 1
                if run_length <= len(ordered):
                    for start2 in range(0, len(ordered) - run_length + 1):
                        window2 = ordered[start2:start2 + run_length]
                        run_bad = model.NewBoolVar(f"hoy_run_{teacher}_{week}_{day}_{start2}")
                        model.Add(sum(teach_slot[(teacher, s)] for s in window2) == run_length).OnlyEnforceIf(run_bad)
                        model.Add(sum(teach_slot[(teacher, s)] for s in window2) <= run_length - 1).OnlyEnforceIf(run_bad.Not())
                        objectives.append(priority * run_bad)

    if objectives:
        model.Minimize(sum(objectives))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_seconds
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)
    status_name = solver.StatusName(status)
    if status_name not in {"OPTIMAL", "FEASIBLE"}:
        return SolveResult("Infeasible", f"No solution found at max {max_teacher_cap} teachers per class. Solver status: {status_name}", [], [], [], diagnostics)

    allocations = []
    for task in task_rows:
        tid = task["TaskID"]
        chosen = None
        for teacher in feasible_teachers_by_task[tid]:
            if solver.Value(x[(tid, teacher)]) == 1:
                chosen = teacher
                break
        row = {
            "Week": _clean(task.get("Week")),
            "Day": _clean(task.get("Day")),
            "Period": _clean(task.get("Period")),
            "SlotID": f"{_clean(task.get('Week'))}_{_clean(task.get('Day'))}_{_clean(task.get('Period'))}",
            "Subject": _clean(class_map[class_by_task[tid]].get("Subject")),
            "Block": _clean(task.get("Block")) or f"{_clean(class_map[class_by_task[tid]].get('Year'))}{_clean(class_map[class_by_task[tid]].get('Side'))}",
            "ClassID": class_by_task[tid],
            "CurrentTeacher": _clean(task.get("CurrentTeacher")),
            "NewTeacher": chosen,
        }
        allocations.append(row)

    teacher_daily = []
    for teacher in teacher_names:
        for week, day in days:
            day_slots = [(week, day, p) for p in teaching_periods if (week, day, p) in all_slots]
            available = [s for s in day_slots if (teacher, s) not in unavailable]
            lessons_count = sum(1 for s in available if solver.Value(teach_slot[(teacher, s)]) == 1)
            role = teacher_role_map.get(teacher, {})
            target_frees = _to_int(role.get("TargetFreesPerDay"), 0) if role else 0
            teacher_daily.append({
                "Teacher": teacher,
                "Week": week,
                "Day": day,
                "AvailablePeriods": len(available),
                "Lessons": lessons_count,
                "Frees": len(available) - lessons_count,
                "FullDay": "Yes" if available and lessons_count == len(available) else "No",
                "Role": _clean(role.get("Role")) if role else "",
                "ProtectedRole": "Yes" if role and _truthy(role.get("Protected", True)) else "No",
                "TargetFreesPerDay": target_frees if role else "",
                "MeetsProtectedFreeTarget": "Yes" if not role or len(available) - lessons_count >= target_frees else "No",
            })

    class_summary = []
    for class_id, tids in sorted(tasks_by_class.items()):
        current_counts = defaultdict(int)
        new_counts = defaultdict(int)
        for tid in tids:
            current = _clean(task_rows[tid].get("CurrentTeacher"))
            if current:
                current_counts[current] += 1
        for allocation in allocations:
            if allocation["ClassID"] == class_id:
                new_counts[allocation["NewTeacher"]] += 1
        def fmt_counts(d: Dict[str, int]) -> str:
            return ", ".join(f"{t} ({n})" for t, n in sorted(d.items(), key=lambda kv: (-kv[1], kv[0])))
        class_summary.append({
            "ClassID": class_id,
            "Block": next((a["Block"] for a in allocations if a["ClassID"] == class_id), ""),
            "Lessons": len(tids),
            "CurrentAllocation": fmt_counts(current_counts),
            "NewAllocation": fmt_counts(new_counts),
            "TeacherCount": len(new_counts),
            "Status": "One teacher" if len(new_counts) == 1 else "Split",
        })

    protected_teachers = [t for t, r in teacher_role_map.items() if _truthy(r.get("Protected", True))]
    if protected_teachers:
        diagnostics.append("Protected role welfare applied to: " + ", ".join(sorted(protected_teachers)))

    return SolveResult(
        status="Optimal" if status_name == "OPTIMAL" else "Feasible",
        message=f"Solved with maximum {max_teacher_cap} teachers per class.",
        allocations=allocations,
        teacher_daily=teacher_daily,
        class_summary=class_summary,
        diagnostics=diagnostics,
        objective=solver.ObjectiveValue(),
    )
