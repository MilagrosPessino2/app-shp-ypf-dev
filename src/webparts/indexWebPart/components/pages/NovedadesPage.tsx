import * as React from 'react';
import styles from './NovedadesPage.module.scss';
import CrearWord from '../crearWord/CrearWord';
import { WebPartContext } from '@microsoft/sp-webpart-base';

type Props = { context: WebPartContext };

const NovedadesPages: React.FC<Props> = ({ context }) => {
    return (
        <div className={styles.page}>
            <header className={styles.header}>
                <div className={styles.headerInner}>
                    <div className={styles.brand}>
                        <img
                             src={require('../../../../public/Logo_de_YPF.png')}
                            alt='YPF'
                            className={styles.logo}
                        />
                    </div>
                    {/* por si querés acciones futuras a la derecha */}
                    <nav className={styles.actions} />
                </div>
            </header>

            <main className={styles.main}>
                <section className={styles.card}>
                    <h1 className={styles.title}>Generador de Novedades</h1>
                    <p className={styles.subtitle}>
                        Exportá un .docx con el formato corporativo.
                    </p>

                    {/* componente existente */}
                    <div className={styles.ctaRow}>
                        <CrearWord context={context} />
                    </div>
                </section>
            </main>

            <footer className={styles.footer}>
                <span>© {new Date().getFullYear()} YPF — Devs : Mili y Cisco</span>
            </footer>
        </div>
    );
};

export default NovedadesPages;
