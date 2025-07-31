import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from "./QuickNotesWebPart.module.scss";
import * as strings from 'QuickNotesWebPartStrings';

import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import quotesdata from './data/quotes.json';

export interface IQuickNotesWebPartProps {
    title: string;
    category: string;
}

export default class QuickNotesWebPart extends BaseClientSideWebPart<IQuickNotesWebPartProps> {

    private _sp: SPFI;

    public async render(): Promise<void> {
        this._sp = spfi().using(SPFx(this.context));
        const listName = "InspirationNotes";
        await this.ensureListExists(listName);

        const quotes: { [key: string]: { quote: string; author: string }[] } = quotesdata;

        const selectedCategory = this.properties.category || "Motivational";
        const validCategories = Object.keys(quotes);
        const safeCategory = validCategories.includes(selectedCategory) ? selectedCategory : "Motivational";
        const randomQuoteList = quotes[safeCategory];
        const randomQuote = randomQuoteList[Math.floor(Math.random() * randomQuoteList.length)];
        const today = new Date().toLocaleDateString();

        this.domElement.innerHTML = `
        <section class="${styles.dailyInspiration}">
            <div class="${styles.card}">
                <h2>${escape(this.properties.title || "Inspiration Card")}</h2>
                <p class="${styles.date}">${today}</p>
                <p class="${styles.quote}">"${randomQuote.quote}"</p>
                <p class="${styles.author}">– ${randomQuote.author}</p>
                <p class="${styles.category}">${selectedCategory}</p>
                <div class="${styles.notes}">
                    <textarea id="noteInput" rows="3" placeholder="Write your personal note here..."></textarea>
                    <button id="saveNoteButton">Save Note to SharePoint</button>
                    <div id="noteList" class="${styles.noteList}"></div>
                </div>
            </div>
        </section>`;

        this.bindSaveHandler(randomQuote.quote, randomQuote.author, selectedCategory, listName);
        await this.displayRecentNotes(listName);
    }

    private async ensureListExists(listName: string): Promise<void> {
        const lists = await this._sp.web.lists();
        const exists = lists.some(list => list.Title === listName);

        if (!exists) {
            await this._sp.web.lists.add(listName, "User inspiration notes", 100);
        }

        const list = this._sp.web.lists.getByTitle(listName);

        const existingFields = await list.fields.select("InternalName")();
        const existingFieldNames = existingFields.map(f => f.InternalName);

        const ensureField = async (internalName: string, schemaXml: string) => {
            if (!existingFieldNames.includes(internalName)) {
                await list.fields.createFieldAsXml(schemaXml);
            }
        };

        await ensureField("Note", `<Field DisplayName="Note" Name="Note" Type="Note" />`);
        await ensureField("Quote", `<Field DisplayName="Quote" Name="Quote" Type="Text" />`);
        await ensureField("Author", `<Field DisplayName="Author" Name="Author" Type="Text" />`);
        await ensureField("Category", `<Field DisplayName="Category" Name="Category" Type="Text" />`);
        await ensureField("Date", `<Field DisplayName="Date" Name="Date" Type="DateTime" Format="DateOnly" />`);

        const viewFields = await list.defaultView.fields();

        const addFieldToView = async (fieldName: string) => {
            if (!viewFields.Items.includes(fieldName)) {
                await list.defaultView.fields.add(fieldName);
            }
        };

        await addFieldToView("Note");
        await addFieldToView("Quote");
        await addFieldToView("Author");
        await addFieldToView("Category");
        await addFieldToView("Date");
    }

    private async displayRecentNotes(listName: string): Promise<void> {
        const items = await this._sp.web.lists.getByTitle(listName).items
            .select("Id", "Title", "Note", "Quote", "Category", "Date")
            .orderBy("Id", false)
            .top(10)();

        const listContainer = this.domElement.querySelector("#noteList") as HTMLElement;
        if (listContainer) {
            listContainer.innerHTML = "<h4>Recent Notes:</h4>" + items.map(item => `
                <div class="${styles.noteItem}" data-id="${item.Id}">
                    <p>"${item.Note}"</p>
                    <p class="${styles.noteMeta}">– ${item.Title}, ${item.Quote}, ${item.Category}, ${new Date(item.Date).toLocaleDateString()}</p>
                    <button class="editNoteButton" data-id="${item.Id}">Edit</button>
                    <button class="deleteNoteButton" data-id="${item.Id}">Delete</button>
                </div>`).join('');
        }
        this.bindEditAndDeleteHandlers(listName);
    }

    private bindSaveHandler(quote: string, author: string, category: string, listName: string): void {
        const saveButton = this.domElement.querySelector("#saveNoteButton") as HTMLButtonElement;
        const noteInput = this.domElement.querySelector("#noteInput") as HTMLTextAreaElement;

        saveButton.onclick = async () => {
            const noteText = noteInput.value;
            if (!noteText.trim()) return;

            try {
                const today = new Date().toISOString();
                await this._sp.web.lists.getByTitle(listName).items.add({
                    Title: this.properties.title,
                    Note: noteText,
                    Quote: quote,
                    Category: category,
                    Date: today
                });
                noteInput.value = "";
                await this.displayRecentNotes(listName);
            } catch (error) {
                console.error("Error saving note:", error);
            }
        };
    }

    private bindEditAndDeleteHandlers(listName: string): void {
        const listContainer = this.domElement.querySelector("#noteList") as HTMLElement;

        listContainer.querySelectorAll(".editNoteButton").forEach(button => {
            button.addEventListener("click", async (e) => {
                const itemId = parseInt((e.target as HTMLElement).getAttribute("data-id") || "0");
                if (!itemId) return;

                const item = await this._sp.web.lists.getByTitle(listName).items.getById(itemId)
                    .select("Note", "Quote", "Category", "Date")();

                const modal = document.createElement("div");
                modal.classList.add(styles.modalOverlay);
                modal.innerHTML = `
  <div class="${styles.modalCard}">
    <h3>Edit Note</h3>
    <div class="${styles.formGroup}">
      <label for="editNote">Note:</label>
      <textarea id="editNote" rows="3">${item.Note || ""}</textarea>
    </div>
    <div class="${styles.formGroup}">
      <label for="editQuote">Quote:</label>
      <input type="text" id="editQuote" value="${item.Quote || ""}" />
    </div>
    <div class="${styles.formGroup}">
      <label for="editCategory">Category:</label>
      <input type="text" id="editCategory" value="${item.Category || ""}" />
    </div>
    <div class="${styles.formGroup}">
      <label for="editDate">Date:</label>
      <input type="date" id="editDate" value="${item.Date ? new Date(item.Date).toISOString().substring(0, 10) : ""}" />
    </div>
    <div class="${styles.modalActions}">
      <button id="saveEdit">💾 Save</button>
      <button id="cancelEdit">❌ Cancel</button>
    </div>
  </div>
`;

                document.body.appendChild(modal);

                modal.querySelector("#saveEdit")?.addEventListener("click", async () => {
                    const updatedNote = (modal.querySelector("#editNote") as HTMLTextAreaElement).value;
                    const updatedQuote = (modal.querySelector("#editQuote") as HTMLInputElement).value;
                    const updatedCategory = (modal.querySelector("#editCategory") as HTMLInputElement).value;
                    const updatedDate = (modal.querySelector("#editDate") as HTMLInputElement).value;

                    try {
                        await this._sp.web.lists.getByTitle(listName).items.getById(itemId).update({
                            Note: updatedNote,
                            Quote: updatedQuote,
                            Category: updatedCategory,
                            Date: updatedDate ? new Date(updatedDate).toISOString() : null
                        });
                        modal.remove();
                        await this.displayRecentNotes(listName);
                    } catch (err) {
                        console.error("Error updating item:", err);
                    }
                });

                modal.querySelector("#cancelEdit")?.addEventListener("click", () => modal.remove());
            });
        });

        listContainer.querySelectorAll(".deleteNoteButton").forEach(button => {
            button.addEventListener("click", async (e) => {
                const itemId = parseInt((e.target as HTMLElement).getAttribute("data-id") || "0");
                if (!itemId) return;
                if (confirm("Are you sure you want to delete this note?")) {
                    try {
                        await this._sp.web.lists.getByTitle(listName).items.getById(itemId).delete();
                        await this.displayRecentNotes(listName);
                    } catch (err) {
                        console.error("Error deleting item:", err);
                    }
                }
            });
        });
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [{
                header: {
                    description: strings.PropertyPaneDescription
                },
                groups: [{
                    groupName: strings.BasicGroupName,
                    groupFields: [
                        PropertyPaneTextField('title', { label: strings.DescriptionFieldLabel }),
                        PropertyPaneDropdown('category', {
                            label: 'Quote Category',
                            options: [
                                { key: 'Motivational', text: 'Motivational' },
                                { key: 'Productivity', text: 'Productivity' },
                                { key: 'Humor', text: 'Humor' }
                            ]
                        })
                    ]
                }]
            }]
        };
    }
}