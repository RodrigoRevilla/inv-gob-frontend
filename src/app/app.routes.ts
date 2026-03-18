import { Routes } from '@angular/router';
import { authGuard } from './guards/auth.guard';
import { adminGuard } from './guards/admin.guard';

export const routes: Routes = [
  { path: '', redirectTo: 'dashboard', pathMatch: 'full' },
  {
    path: 'login',
    loadComponent: () => import('./pages/login/login').then(m => m.LoginComponent),
  },
  {
    path: 'dashboard',
    canActivate: [authGuard],
    loadComponent: () => import('./pages/dashboard/dashboard').then(m => m.DashboardComponent),
  },
  {
    path: 'sesion/:id',
    canActivate: [authGuard],
    loadComponent: () => import('./pages/sesion/sesion').then(m => m.SesionComponent),
  },
  {
    path: 'audit-log',
    canActivate: [authGuard],
    loadComponent: () => import('./pages/audit-log/audit-log').then(m => m.AuditLogComponent),
  },
  {
    path: 'escaneo/:id',
    canActivate: [authGuard],
    loadComponent: () => import('./pages/escaneo/escaneo').then(m => m.EscaneoComponent),
  },
{
  path: 'usuarios',
  canActivate: [authGuard, adminGuard],
  loadComponent: () => import('./pages/usuarios/usuarios').then(m => m.UsuariosComponent),
},
  { path: '**', redirectTo: 'dashboard' },
];