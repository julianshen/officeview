/**
 * The 3-D transition groups, wired (slideshow spec 6.3, S3-7).
 *
 * Each group task exports one registration function (S3-2 geometric, S3-3
 * tiles and particles, S3-4 deformations) and never edits the registry;
 * this file is the one place that calls them. `registerTransitions3D` is
 * idempotent (each group registers the same objects every call), and the
 * show's backend manager calls it before it creates its WebGL2 compositor,
 * so a 3-D row reaches its shader from the first slide. Everything here is
 * in the lazy slideshow chunk; nothing eager imports it.
 *
 * Not a module side effect: the group tests reset the registry and register
 * one group at a time.
 */
import { registerDeform3D } from './deform'
import { registerGeometric3D } from './geometric'
import { registerParticle3D } from './particles'

/** Registers every catalogued 3-D row (all of `THREE_D_IDS`): geometric, tile and particle, deformation. */
export function registerTransitions3D(): void {
  registerGeometric3D()
  registerParticle3D()
  registerDeform3D()
}
