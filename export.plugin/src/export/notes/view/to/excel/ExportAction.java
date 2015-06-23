package export.notes.view.to.excel;

/*
 * Copyright 2015
 * 
 * This file is part of Lotus Notes plugin for export Lotus Notes View into Microsoft Excel.
 * 
 * Licensed under the Apache License, Version 2.0 (the "License"); 
 * you may not use this file except in compliance with the License. 
 * You may obtain a copy of the License at:
 * 
 * http://www.apache.org/licenses/LICENSE-2.0 
 * 
 * Unless required by applicable law or agreed to in writing, software 
 * distributed under the License is distributed on an "AS IS" BASIS, 
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or 
 * implied. See the License for the specific language governing 
 * permissions and limitations under the License.
 */

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import lotus.domino.NotesException;
import lotus.domino.NotesThread;
import lotus.domino.Session;
import lotus.domino.View;
import lotus.domino.ViewEntry;
import lotus.domino.ViewNavigator;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.core.runtime.IStatus;
import org.eclipse.core.runtime.MultiStatus;
import org.eclipse.core.runtime.Status;
import org.eclipse.core.runtime.jobs.IJobChangeEvent;
import org.eclipse.core.runtime.jobs.IJobChangeListener;
import org.eclipse.core.runtime.jobs.Job;
import org.eclipse.jface.action.IAction;
import org.eclipse.jface.dialogs.ErrorDialog;
import org.eclipse.jface.viewers.ISelection;
import org.eclipse.swt.SWT;
import org.eclipse.swt.program.Program;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.ui.IWorkbenchWindow;
import org.eclipse.ui.IWorkbenchWindowActionDelegate;
import org.eclipse.ui.PlatformUI;
import org.eclipse.ui.progress.UIJob;

import com.ibm.notes.java.api.data.NotesViewData;
import com.ibm.notes.java.api.util.NotesSessionJob;
import com.ibm.notes.java.ui.NotesUIElement;
import com.ibm.notes.java.ui.NotesUIWorkspace;
import com.ibm.notes.java.ui.prompt.Prompt;
import com.ibm.notes.java.ui.views.NotesUIView;

/**
 * Our sample action implements workbench action delegate. The action proxy will be created by the
 * workbench and shown in the UI. When the user tries to use the action, this delegate will be
 * created and execution will be delegated to it.
 * 
 * @see IWorkbenchWindowActionDelegate
 */
public class ExportAction implements IWorkbenchWindowActionDelegate, IJobChangeListener {
	private static final Log log = LogFactory.getLog(Activator.PLUGIN_ID);
	// private IWorkbenchWindow window;
	private String filename;

	/**
	 * The constructor.
	 */
	public ExportAction() {
	}

	/**
	 * The action has been activated. The argument of the method represents the 'real' action
	 * sitting in the workbench UI.
	 * 
	 * @see IWorkbenchWindowActionDelegate#run
	 */
	public void run(IAction action) {
		NotesThread.sinitThread();
		try {
			final NotesViewData data = getViewData();
			if (data != null) {
				if (Prompt.YesNo(Messages.ExportAction_0, Messages.ExportAction_2) == 1) {
					filename = getFilename();
					if (filename != null) {
						NotesSessionJob job = new NotesSessionJob(Messages.ExportAction_0) {
							@SuppressWarnings("unchecked")
							@Override
							protected IStatus runInNotesThread(Session session, final IProgressMonitor monitor)
									throws NotesException {
								monitor.beginTask(Messages.ExportAction_8 + data.getName() + Messages.ExportAction_9
										+ filename + "'.", IProgressMonitor.UNKNOWN);//$NON-NLS-1$
								ExcelWriter writer = new ExcelWriter();
								View view = ((NotesViewData) data).open(session);
								// disable auto updating
								view.setAutoUpdate(false);
								writer.createSheet(view.getName().replaceAll("\\\\", "-").trim());//$NON-NLS-1$ //$NON-NLS-2$
								writer.createTableHeader(view);
								// get a ViewNavigator instance for all the view entries
								ViewNavigator nav = view.createViewNav();
								// and set the size of the preloading cache:
								nav.setBufferMaxEntries(500);
								ViewEntry entry = nav.getFirst();
								int row = 0;
								while (entry != null) {
									// check that the user does not press "Cancel"
									if (monitor.isCanceled()) {
										monitor.done();
										return new Status(IStatus.CANCEL, Activator.PLUGIN_ID, Messages.ExportAction_4);
									}
									if (entry.isValid()) {
										if (entry.isDocument()) {
											writer.cerateRow(++row, entry.getColumnValues());
										}
									}
									ViewEntry tmpEntry = nav.getNext();
									entry.recycle();
									entry = tmpEntry;
								}
								nav.recycle();
								if (view != null) {
									view.recycle();
								}
								writer.setAutoSizeColumns();
								monitor.done();
								try {
									FileOutputStream fileOut = new FileOutputStream(filename);
									writer.getWorkbook().write(fileOut);
									fileOut.close();
								} catch (FileNotFoundException e) {
									return new Status(IStatus.ERROR, Activator.PLUGIN_ID, e.getMessage());
								} catch (IOException e) {
									return new Status(IStatus.ERROR, Activator.PLUGIN_ID, e.getMessage());
								}
								return Status.OK_STATUS;
							}
						};
						// Set up the job, then join with it
						job.setPriority(Job.SHORT);
						job.setUser(true);
						job.schedule();
						// Add ourselves as a listener for when the job ends
						job.addJobChangeListener(this);
					}
				}
			}
		} catch (final Exception e) {
			log.fatal(e.getMessage(), e);
			error(e);
		} finally {
			NotesThread.stermThread();
		}
	}

	private String getFilename() {
		FileDialog dialog = new FileDialog(PlatformUI.getWorkbench().getDisplay().getActiveShell(), SWT.SAVE);
		dialog.setText(Messages.ExportAction_3);
		dialog.setFilterExtensions(new String[] { "*.xlsx" }); //$NON-NLS-1$
		return dialog.open();
	}

	private NotesViewData getViewData() throws NotesException {
		NotesUIWorkspace ws = new NotesUIWorkspace();
		NotesUIElement elem = ws.getCurrentElement();
		// See what type of element we have
		if (elem instanceof NotesUIView) {
			return ((NotesUIView) elem).getViewData();
		} else {
			Prompt.Ok(Messages.ExportAction_0, Messages.ExportAction_1);
			return null;
		}
	}

	/**
	 * Selection in the workbench has been changed. We can change the state of the 'real' action
	 * here if we want, but this can only happen after the delegate has been created.
	 * 
	 * @see IWorkbenchWindowActionDelegate#selectionChanged
	 */
	public void selectionChanged(IAction action, ISelection selection) {
	}

	/**
	 * We can use this method to dispose of any system resources we previously allocated.
	 * 
	 * @see IWorkbenchWindowActionDelegate#dispose
	 */
	public void dispose() {
	}

	/**
	 * We will cache window object in order to be able to provide parent shell for the message
	 * dialog.
	 * 
	 * @see IWorkbenchWindowActionDelegate#init
	 */
	public void init(IWorkbenchWindow window) {
		// this.window = window;
	}

	@Override
	public void aboutToRun(IJobChangeEvent arg0) {
		// TODO Auto-generated method stub

	}

	@Override
	public void awake(IJobChangeEvent arg0) {
		// TODO Auto-generated method stub

	}

	@Override
	public void done(IJobChangeEvent arg0) {
		final IStatus status = arg0.getResult();
		if (status == Status.OK_STATUS) {
			File file = new File(filename);
			if (file.exists()) {
				if (System.getProperty("os.name").toLowerCase().indexOf("win") >= 0) { //$NON-NLS-1$ //$NON-NLS-2$				
					if (!Program.launch(filename)) {
						message(Messages.ExportAction_6);
					}
				} else {
					message(Messages.ExportAction_7 + filename);
				}
			}
		} else {
			message(status.getMessage());
		}
	}

	@Override
	public void running(IJobChangeEvent arg0) {
		// TODO Auto-generated method stub

	}

	@Override
	public void scheduled(IJobChangeEvent arg0) {
		// TODO Auto-generated method stub

	}

	@Override
	public void sleeping(IJobChangeEvent arg0) {
		// TODO Auto-generated method stub
	}

	private void message(final String message) {
		UIJob uijob = new UIJob(Messages.ExportAction_5) {
			public IStatus runInUIThread(IProgressMonitor arg0) {
				Prompt.Ok(Messages.ExportAction_0, message);
				return Status.OK_STATUS;
			}
		};
		uijob.schedule();
	}

	private void error(final Throwable t) {
		UIJob uijob = new UIJob(Messages.ExportAction_5) {
			public IStatus runInUIThread(IProgressMonitor arg0) {
				// build the error message and include the current stack trace
				MultiStatus status = createMultiStatus(t.getMessage(), t);
				// show error dialog
				ErrorDialog.openError(PlatformUI.getWorkbench().getDisplay().getActiveShell(),
						Messages.ExportAction_11, t.getLocalizedMessage(), status);

				return Status.OK_STATUS;
			}

			private MultiStatus createMultiStatus(String msg, Throwable t) {

				List<Status> childStatuses = new ArrayList<Status>();
				StackTraceElement[] stackTraces = Thread.currentThread().getStackTrace();

				for (StackTraceElement stackTrace : stackTraces) {
					Status status = new Status(IStatus.ERROR, Activator.PLUGIN_ID, stackTrace.toString());
					childStatuses.add(status);
				}

				MultiStatus ms = new MultiStatus(Activator.PLUGIN_ID, IStatus.ERROR,
						childStatuses.toArray(new Status[] {}), t.toString(), t);
				return ms;
			}

		};
		uijob.schedule();
	}
}