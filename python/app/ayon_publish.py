"""
Mark comps for publishing through AYON's After Effects "render" creator.

When the artist ticks "Publish with AYON", the tool makes sure each comp it renders has an AYON
render instance with the render target "frames" (use existing frames). After the render finishes,
the artist opens the Publisher, which publishes the files at the render queue output path.
Frame ranges come from the render queue item (see ayon_plugins/publish).

Instances are created through AYON's own creator, so the Publisher shows and edits them as usual.
Turn these off in the After Effects creator settings, or the creator will rename comps and change
their frame range and resolution: ``rename_comp_to_product_name`` and ``force_setting_values``.

Only :func:`publish_comps` needs AYON; :func:`plan` is plain Python.
"""
import logging

logger = logging.getLogger(__name__)

CREATOR = "render"
TARGET = "frames"


def instance_comp_id(instance):
    """The comp id a render instance publishes, or None."""
    members = instance.get("members") or []
    try:
        return int(members[0])
    except (IndexError, TypeError, ValueError):
        return None


def plan(instances, comp_ids):
    """
    Split comps into those that already have a render instance and those that need one.

    :param instances: The create context's instances (anything with ``creator_identifier`` and
        ``get``, such as ``CreatedInstance``).
    :param comp_ids: The comps to publish.
    :returns: ``(existing, missing)``: ``{comp_id: [instances]}`` and the comp ids without one,
        in the order given.
    """
    wanted = [int(comp_id) for comp_id in comp_ids]
    existing = {}
    for instance in instances:
        if getattr(instance, "creator_identifier", None) != CREATOR:
            continue
        comp_id = instance_comp_id(instance)
        if comp_id in wanted:
            existing.setdefault(comp_id, []).append(instance)
    missing = [comp_id for comp_id in dict.fromkeys(wanted) if comp_id not in existing]
    return existing, missing


def publish_comps(comp_ids, variant="Main", review=True):
    """
    Give each comp a render instance that publishes existing frames, and save it to the project.

    Existing instances are switched to "frames" and re-activated. New ones are created with AYON's
    render creator for the current folder and task, always with the comp name in the product name.
    Only call this from the GUI thread of the AYON After Effects host.

    :returns: The number of instances created and the number updated.
    """
    from ayon_core.pipeline import registered_host
    from ayon_core.pipeline.create import CreateContext
    from ayon_aftereffects.api import get_stub

    context = CreateContext(registered_host(), discover_publish_plugins=False)
    existing, missing = plan(context.instances, comp_ids)

    for instances in existing.values():
        for instance in instances:
            instance.creator_attributes["render_target"] = TARGET
            instance["active"] = True

    if missing:
        stub = get_stub()
        # The creator works on the selected comps; select exactly ours, then restore.
        previous = [item.id for item in stub.get_selected_items(comps=True, folders=True, footages=True)]
        stub.select_items(missing)
        try:
            context.create(CREATOR, variant, pre_create_data={
                "use_selection": True,
                "use_composition_name": True,
                "render_target": TARGET,
                "mark_for_review": review,
            })
        finally:
            stub.select_items(previous)

    context.save_changes()
    updated = sum(len(instances) for instances in existing.values())
    logger.info("AYON publish: %d render instance(s) created, %d updated" % (len(missing), updated))
    return len(missing), updated
