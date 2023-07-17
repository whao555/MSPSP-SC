import os
from mayavi import mlab
from traits.api import HasTraits, Instance, on_trait_change
from traitsui.api import View, Item
from mayavi.core.ui.api import MayaviScene, MlabSceneModel, SceneEditor


class BulidDam3D(HasTraits):
    def __init__(self, DamSectionDict):
        super().__init__()
        self.Path = '..\DamBody\STL'
        self.DamSectionDict = DamSectionDict
        self.surfaceSet = []
        self.surfaceDic = {}
        self.selectedObj = []
    scene = Instance(MlabSceneModel, ())
    view = View(Item('scene', editor=SceneEditor(scene_class=MayaviScene),
                         height=250, width=300, show_label=False),
                    resizable=True)  # We need this to resize with the parent widget


    @on_trait_change('scene.activated')
    def ReadandBulid(self):
        self.fig = mlab.gcf()
        self.fig.scene.disable_render = True
        for key in self.DamSectionDict.keys():
            OpenPath = self.Path + "\\" + key
            files = os.listdir(OpenPath)
            for file in files:
                file_path = os.path.join(OpenPath, file)
                if os.path.isfile(file_path):
                    StrTem = file.split("-")[0]
                    Elevation = float(StrTem[14:])
                    if Elevation <= float(self.DamSectionDict[key]):
                        poly_data_reader = self.scene.mlab.pipeline.open(file_path)
                        if Elevation == float(self.DamSectionDict[key]):
                            surface = self.scene.mlab.pipeline.surface(poly_data_reader)
                            self.surfaceSet.append(surface.actor.actor._vtk_obj)
                            self.surfaceDic[surface.actor.actor._vtk_obj] = file
                        else:
                            self.scene.mlab.pipeline.surface(poly_data_reader)
        self.fig.scene.disable_render = False
        self.fig.on_mouse_pick(self.picker_callback, type='cell', button='Left')

    def picker_callback(self, picker_obj):
        picked = picker_obj.actors
        # 求交集
        listTem = list(set(self.surfaceSet).intersection(set([o._vtk_obj for o in picked])))
        if picker_obj.actor.property.color == (1.0, 0.0, 0.0):
            picker_obj.actor.property.color = (1.0, 1.0, 1.0)
            self.selectedObj.remove(self.surfaceDic[listTem[0]])
        else:
            if listTem:
                picker_obj.actor.property.color = (1.0, 0.0, 0.0)
                self.selectedObj.append(self.surfaceDic[listTem[0]])

    def selectedObj(self):
        return self.selectedObj

