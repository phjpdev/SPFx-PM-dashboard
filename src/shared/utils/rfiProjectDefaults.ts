import { IProject, IRfi } from '../models/IProject';

/** Copy client contact fields and RFI number from a project onto a new RFI. */
export function applyProjectDefaultsToRfi(
  rfi: IRfi,
  project: IProject,
  rfis: IRfi[],
  opts?: { byCompany?: string }
): IRfi {
  const count = rfis.filter(r => r.projectId === project.id).length;
  const seq = String(count + 1).padStart(3, '0');
  return {
    ...rfi,
    projectId: project.id,
    projectName: project.name,
    rfiNum: `${project.projNum}-RFI-${seq}`,
    submittedTo: project.contact || rfi.submittedTo,
    toCompany: project.company || rfi.toCompany,
    email: project.email || rfi.email || '',
    byCompany: opts?.byCompany ?? rfi.byCompany,
  };
}

/** Resolve project by id and apply defaults when creating a new RFI. */
export function applyProjectDefaultsById(
  rfi: IRfi,
  projectId: string,
  projects: IProject[],
  rfis: IRfi[],
  opts?: { byCompany?: string }
): IRfi {
  const project = projects.find(p => p.id === projectId || p.projNum === projectId);
  if (!project) {
    return { ...rfi, projectId, projectName: rfi.projectName };
  }
  return applyProjectDefaultsToRfi(rfi, project, rfis, opts);
}
