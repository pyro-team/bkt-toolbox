
import bkt


# =================
# = reusing a tab =
# =================

bkt.powerpoint.add_tab(
    bkt.ribbon.Tab(
        label=u'Demo: idQ-Tab',
        id_q='nsBKT:my-unique-id',
        children = [
            bkt.ribbon.Group(
                label=u'Demo: idQ-Group',
                id_q=u'nsBKT:first-group',
                children = [
                    bkt.ribbon.Label(label="Some element")
                ]
            )
        ]
    )
)


bkt.powerpoint.add_tab(
    bkt.ribbon.Tab(
        id_q='nsBKT:my-unique-id',
        children = [
            bkt.ribbon.Group(
                label=u'Next group',
                id_q=u'nsBKT:next-group',
                children = [
                    bkt.ribbon.Label(label="Element in another group")
                ]
            )
        ]
    )
)


# =============================================
# = referencing existing group: insert before =
# =============================================

bkt.powerpoint.add_tab(
    bkt.ribbon.Tab(
        id_q='nsBKT:my-unique-id',
        children = [
            bkt.ribbon.Group(
                insert_before_q=u'nsBKT:first-group',
                label='Another group inserted before',
                children = [
                    bkt.ribbon.Label(label="Another element")
                ]
            )
        ]
    )
)

# bkt.powerpoint.add_tab(
#     bkt.ribbon.Tab(
#         label=u'Demo: idQ-Tab (2)',
#         children = [
#             bkt.ribbon.Group(
#                 id_q=u'nsBKT:first-group',
#                 children=[
#                     bkt.ribbon.Label(label="Another element")
#                 ]
#             )
#         ]
#     )
# )
