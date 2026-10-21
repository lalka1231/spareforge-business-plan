import bpy, math, random
from mathutils import Vector
from pathlib import Path
OUT=Path('/Users/daniilvazdaev/Desktop/уник/Бизнес-планирование/AgroVector/presentations/cinematic_assets');OUT.mkdir(exist_ok=True)
bpy.ops.object.select_all(action='SELECT');bpy.ops.object.delete(use_global=False)
def mat(name,c,metal=0):
 m=bpy.data.materials.new(name);m.diffuse_color=(*c,1);m.use_nodes=True;p=m.node_tree.nodes.get('Principled BSDF');p.inputs['Base Color'].default_value=(*c,1);p.inputs['Metallic'].default_value=metal;p.inputs['Roughness'].default_value=.38;return m
black=mat('carbon',(0.035,.055,.052),.6);silver=mat('steel',(.25,.32,.32),.8);white=mat('polymer',(.8,.84,.75));green=mat('lime',(.48,.68,.12));gold=mat('wheat',(.57,.35,.085));soil=mat('soil',(.15,.10,.035))
def cube(name,loc,scale,m):
 bpy.ops.mesh.primitive_cube_add(size=1,location=loc);o=bpy.context.object;o.name=name;o.scale=scale;o.data.materials.append(m);bpy.ops.object.transform_apply(location=False,rotation=False,scale=True);be=o.modifiers.new('soft edge','BEVEL');be.width=.05;be.segments=3;o.modifiers.new('normals','WEIGHTED_NORMAL');return o
def rod(a,b,r,m):
 d=Vector(b)-Vector(a);bpy.ops.mesh.primitive_cylinder_add(vertices=16,radius=r,depth=d.length,location=(Vector(a)+Vector(b))/2);o=bpy.context.object;o.rotation_euler=d.to_track_quat('Z','Y').to_euler();o.data.materials.append(m);return o
cube('main chassis',(0,0,.2),(.9,.65,.28),black);cube('tank',(0,0,-.18),(.68,.55,.55),white);cube('cover',(0,0,.39),(.65,.49,.12),green)
for x in (-1.2,1.2):
 for y in (-.9,.9):
  rod((x*.2,y*.2,.2),(x,y,.25),.065,black);rod((x,y,.23),(x,y,.43),.1,silver)
  for angle in (0,math.pi/2):
   o=cube('propeller',(x,y,.46),(1.25,.075,.016),black);o.rotation_euler.z=angle
  rod((x*.65,y*.6,.05),(x*.65,y*.6,-.85),.03,silver)
for x in (-.8,.8):rod((x,-.8,-.85),(x,.8,-.85),.035,black)
rod((-1.05,0,-.35),(1.05,0,-.35),.025,silver)
for x in (-.9,-.45,.45,.9):rod((x,0,-.35),(x,0,-.48),.03,green)
scene=bpy.context.scene;scene.render.engine='CYCLES';scene.cycles.samples=24
scene.render.resolution_x=1200;scene.render.resolution_y=750;scene.render.resolution_percentage=100
scene.world.color=(.45,.5,.6)
bpy.ops.object.light_add(type='AREA',location=(2,-3,6));bpy.context.object.data.energy=600;bpy.context.object.data.shape='DISK';bpy.context.object.data.size=5
bpy.ops.object.light_add(type='AREA',location=(-3,3,4));bpy.context.object.data.energy=450;bpy.context.object.data.size=4
bpy.ops.object.camera_add(location=(4,-6,3));cam=bpy.context.object;cam.rotation_euler=(Vector((0,0,0))-cam.location).to_track_quat('-Z','Y').to_euler();cam.data.type='ORTHO';cam.data.ortho_scale=4.8;scene.camera=cam
scene.render.film_transparent=True;scene.render.image_settings.file_format='PNG';scene.render.filepath=str(OUT/'drone_render.png');scene.view_settings.view_transform='AgX';bpy.ops.wm.save_as_mainfile(filepath=str(OUT/'agro_drone.blend'));bpy.ops.render.render(write_still=True)
